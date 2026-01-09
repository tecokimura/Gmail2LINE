# gmail2line（GAS）- Gmail → LINE 通知スクリプト

本リポジトリは、書籍『Gmailに届いたメールをLINEに送信する』で利用する **Google Apps Script（GAS）用ソースコード** を公開するためのものです。  
**このリポジトリ単体で完結する手順は用意していません**。設定手順は書籍をご参照ください。

---

## できること

- Gmail に届いたメールのうち **指定したラベルが付いたメール** だけを定期的に検索
- 件名・本文を整形して **LINE グループへ通知**
- Webhook を受けてスプレッドシートにログ記録（検証・グループID取得用）

---

## 動作イメージ

1. 一定間隔（例：1分ごと）で `executeMain()` が実行される（トリガー）
2. Gmail から「指定ラベル付きメール」を検索する
3. 新規メールがあれば、内容を整形して LINE グループへ送信する

---

## 必要な設定（スクリプトプロパティ）

GAS の **プロジェクト設定 → スクリプトプロパティ** に以下を登録します。

| Key | 内容 |
|---|---|
| `LINE_MESSAPI_TOKEN_MAIN` | LINE Messaging API のチャンネルアクセストークン |
| `LINE_GROUP_ID_MAIN` | 通知を送信する LINE グループID |
| `MAIL_LABEL_NAME` | Gmail の通知対象ラベル名（例：`LINE通知`） |
| `MAIL_GET_INTERVAL_MINUTE` | メールチェック間隔（分）※トリガーと同じ値にする |

---

## 使い方（書籍の流れ）

書籍の手順に沿って進める前提です。

- スクリプトを貼り付け
- スクリプトプロパティを設定
- `initialSetupCheck()` で初期チェック
- 必要なら段階的テスト（`step1_...`〜`step4_...`）
- 最後にトリガーで `executeMain()` を定期実行

---

## 主な関数

- `initialSetupCheck()`  
  初回セットアップ確認（プロパティ / Gmailラベル / LINE API）
- `executeLineMessageTest()`  
  LINE 送信テスト
- `executeMain()`  
  本番の定期実行（Gmail取得 → LINE通知）
- `doPost(e)` / `doGet(e)`  
  Webhook 受信（ログ記録）

---

## Webhook について（重要）

Webhook は **LINE Developers の検証** や **グループID取得** のために使用します。  
**通常の Gmail → LINE 通知だけなら必須ではありません**。セットアップが終わったら OFF にしても問題ありません。  
（ON のままだと検証や通信のたびにスプレッドシートへログが溜まります）

---

## カスタマイズ（任意）

メール本文・件名の整形ルールは、以下の配列を編集することで調整できます。

```js
const SUBJECT_DELETE_PATTERNS = ["WORD1 ", "WORD2 "];
const SUBJECT_DELETE_TO_END_PATTERNS = ["Please add string"];
const BODY_DELETE_PATTERNS = ["NAME1 ", "NAME2 "];
const BODY_DELETE_TO_END_PATTERNS = ["━━━━━━━━━━━━━━━━━━", "******************", "------------------", "=================="];
```

※ 書籍の画面例や出力とズレる可能性があるため、慣れるまでは変更しないのがおすすめです。

---

## 注意事項

- LINE Messaging API（無料プラン）は **月あたりの送信宛先数上限（例：200）** があります  
  ※「200回送れる」ではなく「宛先人数×送信数」で消費されるイメージです
- **トリガーの実行間隔** と **`MAIL_GET_INTERVAL_MINUTE`** は必ず同じ値にしてください
- Gmail / LINE の仕様変更により動かなくなる可能性があります
- 本スクリプトの利用は **自己責任** でお願いします

---

## 免責

本リポジトリのソースコードを使用したことによるトラブル・損害について、作者は責任を負いません。

