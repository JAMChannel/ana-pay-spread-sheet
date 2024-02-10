require 'google_drive'
require 'google/apis/gmail_v1'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'fileutils'
require 'date'
require 'pry'
require 'retryable'
require 'dotenv'
Dotenv.load

# 環境変数からパスを読み込む
CREDENTIALS_PATH = ENV['GCP_CREDENTIALS_PATH']
TOKEN_PATH = ENV['GCP_TOKEN_PATH']
SERVICE_ACCOUNT_KEY = ENV['GCP_SERVICE_ACCOUNT_KEY']
SPREADSHEET_ID = ENV['SPREADSHEET_ID']

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'.freeze
APPLICATION_NAME = 'Gmail API Ruby Quickstart'.freeze
SCOPE = Google::Apis::GmailV1::AUTH_GMAIL_READONLY

class ANAPay
  attr_accessor :email_date, :date_of_use, :amount, :store

  def initialize(email_date: nil, date_of_use: nil, amount: nil, store: nil)
    @email_date = email_date
    @date_of_use = date_of_use
    @amount = amount
    @store = store
  end
end


def authorize
  puts "authorizeメソッドが開始されました。"
  client_id = Google::Auth::ClientId.from_file CREDENTIALS_PATH
  token_store = Google::Auth::Stores::FileTokenStore.new file: TOKEN_PATH
  authorizer = Google::Auth::UserAuthorizer.new client_id, SCOPE, token_store
  user_id = 'default'
  credentials = authorizer.get_credentials user_id
  if credentials.nil?
    url = authorizer.get_authorization_url base_url: OOB_URI
    puts 'Open the following URL in the browser and enter the ' \
         "resulting code after authorization:\n" + url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI
    )
  end
  puts "authorizeメソッドが完了しました。"
  credentials
end

def create_session
  puts "create_sessionメソッドが開始されました。"
  # サービスアカウントのJSONキーファイルパス
  session = GoogleDrive::Session.from_service_account_key(SERVICE_ACCOUNT_KEY)
  return session
  puts "create_sessionメソッドが完了しました。"
end

# スプレッドシートにANA Payの情報を書き込む
def write_to_spreadsheet(session, ana_pay)
  puts "write_to_spreadsheetメソッドが開始されました。"
  # スプレッドシートIDを指定
  spreadsheet_id = SPREADSHEET_ID
  # スプレッドシートを開く
  begin
    spreadsheet = session.spreadsheet_by_key(spreadsheet_id)
    # 最初のワークシートを選択
    worksheet = spreadsheet.worksheets.first

    # 最後の行を見つける
    last_row = worksheet.num_rows + 1

    # データを追加する行を指定して、値を設定
    worksheet[last_row, 1] = ana_pay.email_date.strftime('%Y/%m/%d %H:%M:%S') unless ana_pay.email_date.nil?
    worksheet[last_row, 2] = ana_pay.date_of_use.strftime('%Y/%m/%d %H:%M:%S') unless ana_pay.date_of_use.nil?
    worksheet[last_row, 3] = ana_pay.amount unless ana_pay.amount.nil?
    worksheet[last_row, 4] = ana_pay.store unless ana_pay.store.nil?

    # スプレッドシートを保存
    worksheet.save
    puts "write_to_spreadsheetメソッドが完了しました。"
  rescue Google::Apis::RateLimitError => e
    puts "Rate limit exceeded, retrying in 60 seconds..."
    sleep(60)
    retry
  end
end

# ANA Payの情報をメールから取得する
def get_anapay_info_from_email(service)
  puts "get_anapay_info_from_emailメソッドが開始されました。"
  user_id = 'me'
  query = 'from:payinfo@121.ana.co.jp subject:ご利用のお知らせ after:2023/1/1 before:2023/12/31'
  session = create_session

  # メッセージを取得するためのループを開始
  next_page_token = nil
  begin
    response = service.list_user_messages(user_id, q: query, page_token: next_page_token)
    response.messages&.each do |message|
      msg = service.get_user_message(user_id, message.id)
      ana_pay = parse_message_for_anapay_info(msg)
      # 得られた情報をスプレッドシートに書き込む
      Retryable.retryable(tries: 3, on: [Google::Apis::ServerError, Google::Apis::ClientError]) do
        write_to_spreadsheet(session, ana_pay)
        # 1秒待機
        sleep 1
      end
    end
    next_page_token = response.next_page_token
  end while next_page_token

  puts "get_anapay_info_from_emailメソッドが完了しました。"
end

def parse_message_for_anapay_info(msg)
  ana_pay = ANAPay.new
  msg.payload.headers.each do |header|
    case header.name
    when "Date"
      date_str = header.value.gsub(" +0900 (JST)", "")
      ana_pay.email_date = DateTime.parse(date_str)
    end
  end

  body_data = if msg.payload.parts
                part = msg.payload.parts.find { |p| p.mime_type == 'text/plain' }
                part&.body&.data
              else
                msg.payload.body.data
              end
  if body_data
    body = body_data.force_encoding('UTF-8')

    body.each_line do |line|
      if line.start_with?("ご利用")
        key, value = line.split("：")
        case key
        when "ご利用日時"
          ana_pay.date_of_use = DateTime.parse(value)
        when "ご利用金額"
          ana_pay.amount = value.gsub(",", "").gsub("円", "").to_i
        when "ご利用店舗"
          ana_pay.store = value.strip
        end
      end
    end
  end

  ana_pay
rescue ArgumentError => e
  puts "ArgumentError: #{e.message}"
end

# Initialize the API
def main
  service = Google::Apis::GmailV1::GmailService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize


  puts "ANA Payの情報をメールから取得します。"

  # ANA Payの情報をメールから取得する
  anapay_info = get_anapay_info_from_email(service)

  puts "ANA Payの情報の取得が完了しました。"

  puts "mainメソッドが終了しました。"
end

main if __FILE__ == $0
