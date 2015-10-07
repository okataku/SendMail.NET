# SendMail

## 概要
メール送信機能のみを実装したコンソールアプリケーションです。  
本アプリケーションでメールを送信する為には、STMPサーバーの情報と送信メッセージの内容を指定する必要があります。コマンドライン引数で指定する方法、XMLファイルで定義する方法、前2つの方法を組み合わせる方法の3パターンの指定方法があります。  
メール送信処理は.Net Frameworkの`System.Net.Mail.SmtpClient`に依存しています。

## コマンドライン引数
### SMTPサーバー情報
|オプション|説明|
|---|---|
|`/H host`|サーバー名またはIPを指定する。|
|`/P port`|ポート番号を指定する。|
|`/S`|SSLを使用する場合は指定する。|
|`/U user:passwd `|認証情報を指定する。|
|`/T 6000`|送信タイムアウト時間をミリ秒で指定する。|

### 送信メッセージ内容
|オプション|説明|
|---|---|
|`/e charset`|件名の文字エンコーディングを指定する。|
|`/s subject`|件名を指定する。|
|`/f foo@var.com<foo>`|送信元メールアドレスを指定する。`<foo>`は表示名。省略可能。|
|`/t foopvar.com<foo>`|宛先(To)を指定する。`<foo>`は表示名。省略可能。宛先を複数指定する場合は`;`で連結する。|
|`/c foopvar.com<foo>`|宛先(Cc)を指定する。`<foo>`は表示名。省略可能。宛先を複数指定する場合は`;`で連結する。|
|`/b foopvar.com<foo>`|宛先(Cc)を指定する。`<foo>`は表示名。省略可能。宛先を複数指定する場合は`;`で連結する。|
|`/h`|HTMLメールとして送信する場合は指定する。|
|`/E charset`|本文の文字エンコーディングを指定する。|
|`/B body`|本文を指定する。|
|`/a path\to\file`|添付ファイルを指定する。複数指定する場合は`;`で連結する。|
|`/l path\to\log`|エラーログの出力先を指定する。|

※`/HPS`のような指定はできません。

## XMLファイルで送信内容を指定する方法
XMLファイルにSMTP情報と送信メッセージの内容を定義しておき、実行時に読み込ませる方法です。  
コマンドライン引数として`/x`を指定することでXMLファイルを読み込ませることができます。

### コマンド
```
> SendMail.exe /x path\to\SendMail.exe
```
### XML
以下のXMLを参考に必要な箇所を編集します。
```
<?xml version="1.0" encoding="utf-8" ?>
<sendmail>
  <smtp>
    <host>hostname or ip</host>
    <port>25</port>
    <ssl>false</ssl>
    <user></user>
    <password></password>
    <timeout>60000</timeout>
  </smtp>
  <message>
    <subject encoding="utf-8">Subject</subject>
    <from>
      <address displayName="">foo@var.com</address>
    </from>
    <to>
      <address displayName="">foo@var.com</address>
    </to>
    <cc>
      <address displayName="">foo@var.com</address>
    </cc>
    <bcc>
      <address displayName="">foo@var.com</address>
    </bcc>
    <body encoding="utf-8" html="true">
      Body.
    </body>
    <attachment>
      <file>path/to/file</file>
    </attachment>
  </message>
</sendmail>
```

## XMLとコマンド引数の両方を指定して実行
この場合、内部処理的には、最初にXMLファイルに記述された定義内容が読み込まれます。
その後、コマンドライン引数で指定した値を読み込ますが、**XMLで定義した値はコマンドライン引数の値によって上書きされます**。  
[宛先]のように複数指定できるものに関しては、XMLで指定した複数の宛先すべてが無視され、コマンドライン引数の値が使用されます。例えば、XMLで3つの宛先を指定していても、コマンドライン引数で1つの宛先のみを指定すると、コマンドライン引数で指定した1つの宛先にのみ送信されます。

## 件名・本文に外部から指定した値を挿入
件名と本文に`/v`オプションで指定した値を挿入することが出来ます。  

### 挿入先の指定方法
挿入する場所に`{n}`と記述します(`n`には`/v`で指定した値リストのインデックスを記述)。
```
<subject>SUBJECT {0}</subject>
<body>
<html>
  <head>
    <style>
    .foo {{ color: #fff; }}
    </style>
  </head>
  <body>
    DATE: {0}
    CUSTOMER CODE: {1}
    CUSTOMER NAME: {2}
  </body>
</html>
</body>
```
`{}`の文字を挿入先指定以外の用途で使用する場合は、`{{`、`}}`のようにエスケープする必要があります。  

### 値リストの受け渡し
挿入先に渡す値を`;`区切りで記述する。値に`;`を含めたい場合は`\`でエスケープします。
```
> path\to\SendMail.exe -x \path\to\MailMessage.xml -v "2015/01/01;9999;XXXX"
```

### 挿入処理の結果
```
// Subject
SUBJECT 2015/01/01

// Body
DATE: 2015/01/01
CUSTOMER CODE: 9999
CUSTOMER NAME: XXXX
```
※.NET Frameworkの`System.String.Format`メソッドを使用しています。

## アプリケーションの終了コード
|コード|意味|
|---|---|
|0|正常終了|
|1|異常終了|
