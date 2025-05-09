> ⚠️ **ご注意**  
> Moneyforwardの二段階認証導入により、本ツールは現在正常に利用できない可能性があります。

---

# anapay2moneyforward

* ANA Payのメールから支払い情報を取り出して、マネーフォワードに登録するスクリプト。

## 環境構築

### GmailでIMAPを有効に
 
1. **Gmailにログイン**
   - 設定からIMAPを有効にしてください。

### Moneyforwadで「未対応の電子マネー・プリペイド」から"ANA Pay"を追加

1. **マネーフォワード MEにログイン**
   - 未対応の電子マネー・プリペイドの設定画面に移動し、"ANA Pay"を追加してください。
   
### Google SpreadsheetのAPIを使用する準備

1. **Google Cloudコンソールでプロジェクトを作成**
   - Google Cloudコンソールにアクセスし、新しいプロジェクトを作成します。

2. **Google Drive APIとGoogle Sheets APIを有効にする**
   - プロジェクトでGoogle Drive APIとGoogle Sheets APIを有効にします。

3. **サービスアカウントを作成**
   - プロジェクト内でサービスアカウントを作成し、必要な権限を付与します。
   - サービスアカウントの認証情報をJSON形式でダウンロードし、プロジェクトのルートディレクトリに`service-account.json`という名前で保存します。

### .envファイルの設定

スクリプトがGmail、Google Sheets、およびマネーフォワード MEアカウントにアクセスできるように、以下のように`.env`ファイルを作成します。

\`\`\`
SHEET_ID=YOUR_SHEET_ID
EMAIL=YOUR_MONEYFORWARD_EMAIL
PASSWORD=YOUR_MONEYFORWARD_PASSWORD
GOOGLE_APPLICATION_CREDENTIALS=/path/to/your/service-account.json
\`\`\`

- `SHEET_ID`: Google SpreadsheetのシートIDを設定します。
- `EMAIL`: マネーフォワード MEのログインID（メールアドレス）を設定します。
- `PASSWORD`: マネーフォワード MEのログインパスワードを設定します。
- `GOOGLE_APPLICATION_CREDENTIALS`: サービスアカウントの認証情報ファイルのパスを指定します。

### スクリプトの実行

環境設定が完了したら、以下の手順でスクリプトを実行します。

#### Dockerイメージのビルド

\`\`\`bash
cd /path/to/ANAPayToMoneyForward
docker build -t anapay2moneyforward .
\`\`\`

#### Dockerコンテナの実行

環境変数や認証情報を含むディレクトリをマウントして、Dockerコンテナを実行します。

\`\`\`bash
docker run -d \
    -v /path/to/local/screenshots:/app/screenshots \
    --env-file /path/to/your/.env \
    --name anapay2moneyforward \
    anapay2moneyforward
\`\`\`

#### Dockerコンテナの開始

コンテナを実行するには、以下のコマンドを使用します。

\`\`\`bash
docker start anapay2moneyforward
\`\`\`

これで、ANA Payのメールから支払い情報を抽出し、マネーフォワードに自動登録するプロセスが開始されます。

## オリジナル
このプロジェクトはhttps://github.com/takanory/anapay2moneyforwardを元にカスタマイズしたものです。

## 変更点
- **通常版Moneyforward MEに対応**
- **スクリーンショットの保存機能を追加**
- **.envファイルのサポートを追加**
- **一部処理をHeliumからSeleniumに置き換え**
- **支払元で"ANA Pay"を選択する処理を追加**
- **無料版Gmailでも使いやすいようにGmail APIからIMAPに変更**


## ライセンス
このプロジェクトはオリジナルのリポジトリを基にしており、MITライセンスの下で公開されています。

---
