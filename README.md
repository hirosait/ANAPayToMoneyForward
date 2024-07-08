# anapay2moneyforward

* ANA Payのメールから支払い情報を取り出して、マネーフォワードに登録するスクリプト。

## 環境構築

## Moneyforwadで「未対応の電子マネー・プリペイド」から"ANA Pay"を追加しておく

## GmailとGoogle SpreadsheetのAPIを使えるようにする

* 以下のページも参考にして、GoogleのAPIを使えるようにする
  * [Python クイックスタート  |  Gmail  |  Google for Developers](https://developers.google.com/gmail/api/quickstart/python?hl=ja)
* Google Cloudコンソールでプロジェクトを作成する
  * プロジェクトでGmail APIとGoogle Sheets APIを有効にする
  * OAuth 同意画面でアプリを作成する
  * テストユーザーで自分のGoogleアカウントを追加
  * 認証情報をダウンロードし、`credentials.json` として保存
* 以下のように `quickstart.py` を実行する
  * 自分のGoogleアカウントで **同意する**
  * 処理が成功すると `token.json` が生成される
* quickstart.pyでRefresherrorになる場合
  * 再度実行してもエラーになる場合は、token.jsonを削除して、gcloudコンソールのAPIとサービス→認証情報を作成して再度実行すれば良い

```
(env) $ python quickstart.py
(env) $ ls *.json
credentials.json	token.json
```


## .envファイルの設定
スクリプトがgmail, Google Sheetsm, MoneyForward MEアカウントにアクセスするために、以下のように.envファイルを作成します。

```
SHEET_ID=YOUR_SHEET_ID
EMAIL=<MoneyForwardのID>
PASSWORD=<MoneyForwardのPass>k
```

## スクリプトの実行
環境設定が完了したら、以下のコマンドでスクリプトを実行します。

* Dockerイメージをビルドして実行する

```
cd /path/to/ANAPayToMoneyForward
docker build -t anapay2moneyforward .
```

* Dockerコンテナの実行
  * 環境変数や認証情報を含むディレクトリをマウントして、Dockerコンテナを実行します。

```
docker run -d \
    -v /path/to/local/screenshots:/app/screenshots \
    --name anapay2moneyforward \
    anapay2moneyforward

docker start anapay2moneyforward
```

以上で、ANA Payのメールから支払い情報を取り出し、マネーフォワードに自動登録することができます。

## オリジナル
このプロジェクトはhttps://github.com/takanory/anapay2moneyforwardを元にカスタマイズしたものです。

## 変更点
通常版Moneyforward MEに対応  
スクリーンショットを追加  
.env追加  
heliumからseleniumに一部書き換え（途中）
支払元で”ANA Pay"を選択するように追加

## ライセンス
このプロジェクトはオリジナルのリポジトリを基にしており、MITライセンスの下で公開されています。

---
