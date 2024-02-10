## ana-pay-spread-sheet

- ANAペイがマネーフォワードに連携機能ないからgmailからスプレットシートに書き出し
- 確定申告用
- 数百件しかないから並列やメモリ管理は特にしてない
- APIのレートリミットとクォータの管理は後でメンテしても良さそう
- spreadに関しては毎回saveしてしまっているのでこちらは改善しといたほうがよさそうなので来年使う機会があれば修正

## credentials.jsonの参考
- https://developers.google.com/gmail/api/quickstart/go?hl=ja

```
ruby quickstart.rb 
```

It would have been quicker to type in the number of entries since there were only about 300 or so....
