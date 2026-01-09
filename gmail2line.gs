/**
 * Gmail → LINE 通知 自動化スクリプト
 *
 * このスクリプトは、Gmail に届いたメールのうち「指定したラベルが付いたメール」だけを
 * 定期的にチェックし、その内容を LINE グループへ自動で通知するためのものです。
 * 本書『Gmailに届いたメールをLINEに送信する』の手順に沿って設定・利用することを前提にしています。
 *
 * ※ コードの内容を理解する必要はありません。
 *    設定値（スクリプトプロパティ）だけ正しく入力してください。
 *
 * ─────────────────────────────
 * 主な動作の流れ
 * ─────────────────────────────
 * 1. 一定間隔（例：1分ごと）でスクリプトが実行される
 * 2. Gmailから「指定したラベル付きメール」を検索する
 * 3. 新しく届いたメールがあれば、件名・本文を整形して LINE グループへ送信する
 *
 * ─────────────────────────────
 * 必要な設定（スクリプトプロパティ）
 * ─────────────────────────────
 * ・LINE_MESSAPI_TOKEN_MAIN
 *   LINE Messaging API のチャンネルアクセストークン
 * ・LINE_GROUP_ID_MAIN
 *   通知を送信する LINE グループのID
 * ・MAIL_LABEL_NAME
 *   Gmailで作成した通知対象メール用ラベル名（例: LINE通知）
 * ・MAIL_GET_INTERVAL_MINUTE
 *   メールをチェックする間隔（分）※トリガーと同じ値にしてください
 *
 * ─────────────────────────────
 * 初期セットアップ
 * ─────────────────────────────
 * 1. スクリプトプロパティをすべて設定
 * 2. initialSetupCheck() を実行して設定確認
 * 3. 問題なければ executeMain() をトリガーで定期実行
 *
 * ※ initialSetupCheck などのテスト用関数は初期確認用です。
 *
 * ─────────────────────────────
 * Webhookについて（補足）
 * ─────────────────────────────
 * Webhook は LINE Developers の検証やグループID取得のために使用します。
 * 通常の Gmail → LINE 通知動作には必須ではなく、セットアップ後はオフにしても問題ありません。
 *
 * ─────────────────────────────
 * 注意事項
 * ─────────────────────────────
 * ・LINE Messaging API（無料プラン）には月あたりの送信宛先数上限（200）があります
 * ・トリガーの実行間隔と MAIL_GET_INTERVAL_MINUTE は必ず同じにしてください
 * ・本スクリプトの利用は自己責任でお願いします
 */

// ==========================================
// 基本設定
// ==========================================

// メール取得間隔（分） - 環境変数から取得、トリガー実行間隔と一致させる必要があります
var emailFetchIntervalMinutes = Number(PropertiesService.getScriptProperties().getProperty('MAIL_GET_INTERVAL_MINUTE')); 

// LINE Message API エンドポイント
const LINE_API_PUSH_URL = 'https://api.line.me/v2/bot/message/push';  
const LINE_API_QUOTA_URL = 'https://api.line.me/v2/bot/message/quota/consumption';  

// LINE API認証情報 - 環境変数から自動取得
const mainAccessToken = PropertiesService.getScriptProperties().getProperty('LINE_MESSAPI_TOKEN_MAIN');
const mainGroupId = PropertiesService.getScriptProperties().getProperty('LINE_GROUP_ID_MAIN');

// 複数のLINE APIトークン・グループ設定（拡張用）
// 1つのトークンで送信上限に達した場合に別のトークンを使用する
const lineApiTokenGroups = [
  {
    "token": mainAccessToken,
    "group_id": mainGroupId
  },
  // 追加のトークンを使用する場合はここに設定を追加
  // {
  //   "token": "追加のアクセストークン",
  //   "group_id": "追加のグループID"
  // }
]

// ==========================================
// メール処理設定 - 用途に応じてカスタマイズ可能
// ==========================================

// Gmail検索対象ラベル名 - 通知対象メールに付けるラベル（お使いの環境に合わせて変更してください）
const GMAIL_SEARCH_LABEL = PropertiesService.getScriptProperties().getProperty('MAIL_LABEL_NAME');

// メール件名テキスト処理設定
// 件名から削除したい文字列パターン
const SUBJECT_DELETE_PATTERNS = [
  "WORD1 ",     // 削除したい文字列を設定（例: 会社名など）
  "WORD2 ",     // 削除したい文字列を設定
];

// 件名でこの文字列以降をすべて削除するパターン
const SUBJECT_DELETE_TO_END_PATTERNS = [
  "Please add string",  // 削除開始文字列を設定
];

// メール本文テキスト処理設定
// 本文から削除したい文字列パターン
const BODY_DELETE_PATTERNS = [
  "NAME1 ",     // 削除したい文字列を設定（例: 署名など）
  "NAME2 ",     // 削除したい文字列を設定
];

// 本文でこの文字列以降をすべて削除するパターン
const BODY_DELETE_TO_END_PATTERNS = [
    "━━━━━━━━━━━━━━━━━━",    // 罫線区切り
    "******************",      // アスタリスク区切り
    "------------------",      // ハイフン区切り
    "==================",      // イコール区切り
];

// ==========================================
// システム定数
// ==========================================

// LINE Message APIの制限値
const LINE_API_MAX_MESSAGES_PER_MONTH = 200;

// ==========================================
// 設定検証・ユーティリティ関数
// ==========================================

/**
 * 環境変数の設定チェック
 * スクリプト実行前に必要な設定が正しく行われているかを検証
 * @return {Object} チェック結果 {isValid: boolean, errors: Array}
 */
function validateConfiguration() {
  const errors = [];
  
  if (!mainAccessToken || mainAccessToken.length === 0) {
    errors.push("LINE_MESSAPI_TOKEN_MAIN が設定されていません。スクリプトプロパティに設定してください。");
  }
  
  if (!mainGroupId || mainGroupId.length === 0) {
    errors.push("LINE_GROUP_ID_MAIN が設定されていません。スクリプトプロパティに設定してください。");
  }
  
  if (!emailFetchIntervalMinutes || isNaN(emailFetchIntervalMinutes) || emailFetchIntervalMinutes <= 0) {
    errors.push("MAIL_GET_INTERVAL_MINUTE が正しく設定されていません。正の数値を設定してください。");
  }

  if (!GMAIL_SEARCH_LABEL || GMAIL_SEARCH_LABEL.length === 0) {
    errors.push("MAIL_LABEL_NAME（Gmailラベル名）が設定されていません。スクリプトプロパティに設定してください。");
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * LINE Message API用HTTPヘッダー生成
 * 認証ヘッダーを含むHTTPリクエスト用のヘッダーを作成
 * @param {string} token - LINE APIアクセストークン
 * @return {Object} HTTPヘッダーオブジェクト
 */
function createLineApiHeaders(token) {
  return {
    'Content-Type': 'application/json; charset=UTF-8',
    'Authorization': 'Bearer ' + token
  };
}

/**
 * エラーログ出力関数
 * エラー発生時に詳細な情報をログに記録
 * @param {string} functionName - エラーが発生した関数名
 * @param {string} errorMessage - エラー内容の説明
 * @param {Error} error - エラーオブジェクト（オプション）
 */
function logError(functionName, errorMessage, error = null) {
  const timestamp = new Date().toLocaleString('ja-JP');
  let logMessage = `[エラー] ${timestamp} - ${functionName}: ${errorMessage}`;
  
  if (error) {
    logMessage += `詳細: ${error.toString()}`;
    if (error.stack) {
      logMessage += `スタックトレース: ${error.stack}`;
    }
  }
  
  Logger.log(logMessage);
  console.error(logMessage);
}

/**
 * 情報ログ出力関数
 * 処理状況や結果の情報をログに記録
 * @param {string} message - ログメッセージ
 */
function logInfo(message) {
  const timestamp = new Date().toLocaleString('ja-JP');
  const logMessage = `[情報] ${timestamp} - ${message}`;
  Logger.log(logMessage);
}

/**
 * 警告ログ出力関数
 * 注意が必要な状況をログに記録
 * @param {string} message - 警告メッセージ
 */
function logWarning(message) {
  const timestamp = new Date().toLocaleString('ja-JP');
  const logMessage = `[警告] ${timestamp} - ${message}`;
  Logger.log(logMessage);
}

/**
 * 成功ログ出力関数
 * 処理の成功をログに記録
 * @param {string} message - 成功メッセージ
 */
function logSuccess(message) {
  const timestamp = new Date().toLocaleString('ja-JP');
  const logMessage = `[成功] ${timestamp} - ${message}`;
  Logger.log(logMessage);
}

// ==========================================
// Webhook機能
// ==========================================

// ==========================================
// Webhook機能（GET/POST両対応）
// ==========================================

/**
 * Webhookボディ(JSON文字列)から LINE の groupId を抽出
 * - LINEのwebhook形式の場合のみ groupId を返す
 * - その他の場合やパース失敗時は 空文字("") を返す
 * @param {string} bodyString - e.postData.contents の文字列
 * @return {string} groupId または 空文字
 */
function extractGroupIdFromWebhookBody(bodyString) {
  if (!bodyString) {
    return "";
  }

  try {
    const data = JSON.parse(bodyString);

    // events が配列でなければ終了
    if (!data.events || !Array.isArray(data.events)) {
      return "";
    }

    // 複数イベントが来る可能性も想定してループ
    for (let i = 0; i < data.events.length; i++) {
      const ev = data.events[i];
      if (
        ev &&
        ev.source &&
        ev.source.type === 'group' &&
        ev.source.groupId
      ) {
        return String(ev.source.groupId);
      }
    }

    // 見つからなければ空文字
    return "";
  } catch (error) {
    // JSONパース失敗などはログだけ残して空文字返却
    logError('extractGroupIdFromWebhookBody', 'Webhookボディの解析に失敗しました', error);
    return "";
  }
}

/**
 * 共通のWebhookデータ処理
 * GETとPOSTリクエストの共通処理ロジック
 * @param {Object} requestData - リクエストパラメータ
 * @param {string} requestBody - リクエストボディ
 * @return {Object} レスポンス
 */
function processWebhookRequest(requestData, requestBody) {
  try {
    logInfo('Webhookリクエストを受信しました');
    
    // スプレッドシートにデータを記録
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    
    // タイムスタンプ、パラメータ、ボディを記録
    const timestamp = new Date();
    const parameterString = JSON.stringify(requestData);
    const bodyString = requestBody;
    
    // 既存の通常ログ行（そのまま）
    sheet.appendRow([timestamp, parameterString, bodyString]);

    // body から groupId を抽出して、別行として追記
    const groupId = extractGroupIdFromWebhookBody(bodyString);
    if (groupId) {
      const groupIdInfo = `groupId=${groupId}`;
      sheet.appendRow([timestamp, groupIdInfo, ""]);
      logInfo('groupId を含むログ行を追加しました: ' + groupIdInfo);
    }
    
    logInfo('リクエストをスプレッドシートに記録しました');
    
    // 成功レスポンス
    return ContentService.createTextOutput("OK")
      .setMimeType(ContentService.MimeType.TEXT);
      
  } catch (error) {
    logError('processWebhookRequest', 'リクエスト処理中にエラーが発生しました', error);
    
    // エラーレスポンス
    return ContentService.createTextOutput("ERROR")
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Webhook GET リクエスト受信処理
 * 外部からのGETリクエストを受け取り、スプレッドシートに記録
 * @param {Object} e - GETリクエストのイベントオブジェクト
 * @return {Object} レスポンス
 */
function doGet(e) {
  const requestData = e.parameter || {};
  const requestBody = '';
  
  return processWebhookRequest(requestData, requestBody);
}

/**
 * Webhook POST リクエスト受信処理
 * 外部からのPOSTリクエストを受け取り、スプレッドシートに記録
 * @param {Object} e - POSTリクエストのイベントオブジェクト
 * @return {Object} レスポンス
 */
function doPost(e) {
  const requestData = e.parameter || {};
  const requestBody = e.postData ? e.postData.contents : '';
  
  return processWebhookRequest(requestData, requestBody);
}

// ==========================================
// 初期設定チェック機能
// ==========================================

/**
 * 初回セットアップ用チェック関数
 * スクリプト使用前に必要な設定がすべて整っているかを確認します
 * @return {Object} セットアップ状況 {isComplete: boolean, completedSteps: number, totalSteps: number, status: Object, errors: Array}
 */
function initialSetupCheck() {
  logInfo('=== 初回セットアップチェックを開始します ===');
  
  const setupStatus = {
    scriptProperties: false,
    gmailLabels: false,
    messagingApi: false
  };
  
  const setupErrors = [];
  
  // 1. スクリプトプロパティの確認
  logInfo('STEP 1/3: スクリプトプロパティの確認');
  const configCheck = validateConfiguration();
  
  if (configCheck.isValid) {
    setupStatus.scriptProperties = true;
    logSuccess('スクリプトプロパティは正しく設定されています');
  } else {
    setupErrors.push('スクリプトプロパティの設定が不完全です');
  }
  
  // 2. Gmailラベルの確認
  logInfo('STEP 2/3: Gmailラベルの確認');
  try {
    const labels = GmailApp.getUserLabels();
    const targetLabel = labels.find(label => label.getName() === GMAIL_SEARCH_LABEL);
    
    if (targetLabel) {
      setupStatus.gmailLabels = true;
      logSuccess('ラベルを見つけました');
    } else {
      setupErrors.push(`Gmail ラベル「${GMAIL_SEARCH_LABEL}」が見つかりません`);
    }
  } catch (error) {
    setupErrors.push('Gmailへのアクセスでエラーが発生しました');
    logError('initialSetupCheck', 'Gmail確認中にエラー', error);
  }
  
  // 3. LINE Messaging API設定の確認
  logInfo('STEP 3/3: LINE Messaging API設定の確認');
  if (setupStatus.scriptProperties) {
    try {
      const remaining = getRemainingMessageCount(mainAccessToken);
      if (remaining > 0) {
        setupStatus.messagingApi = true;
        logSuccess(`LINE Messaging API設定は正常です（残り送信数: ${remaining}通）`);
      } else {
        setupErrors.push('LINE Messaging APIの認証に失敗しました');
      }
    } catch (error) {
      setupErrors.push('LINE Messaging API接続でエラーが発生しました');
      logError('initialSetupCheck', 'LINE Messaging API確認中にエラー', error);
    }
  } else {
    Logger.log('スクリプトプロパティを設定してからLINE Messaging API設定を確認します');
  }
  
  // 結果サマリー
  const totalSteps = Object.keys(setupStatus).length;
  const completedSteps = Object.values(setupStatus).filter(status => status).length;
  
  Logger.log('');
  Logger.log('='.repeat(50));
  Logger.log('セットアップ状況サマリー');
  Logger.log('='.repeat(50));
  Logger.log(`進捗: ${completedSteps}/${totalSteps} 完了`);
  
  Logger.log('');
  Logger.log('完了済み:');
  if (setupStatus.scriptProperties) Logger.log('  - スクリプトプロパティ設定');
  if (setupStatus.gmailLabels) Logger.log('  - Gmail ラベル作成');
  if (setupStatus.messagingApi) Logger.log('  - LINE Messaging API 設定');
  
  if (setupErrors.length > 0) {
    Logger.log('');
    Logger.log('未完了・要対応:');
    setupErrors.forEach(error => Logger.log(`  - ${error}`));
  }
  
  // 次のステップの案内
  if (completedSteps === totalSteps) {
    logSuccess('初期セットアップがすべて完了しました');
    Logger.log('');
    Logger.log('次のステップ:');
    Logger.log('1. step1_validateSetup() で詳細な設定確認');
    Logger.log('2. 段階的テスト（step2〜step4）を実行');
    Logger.log('3. 問題がなければトリガーを設定して本番運用開始');
  } else {
    Logger.log('');
    Logger.log('次のステップ:');
    Logger.log('上記の未完了項目を対応してから、もう一度 initialSetupCheck() を実行してください');
  }
  
  return {
    isComplete: completedSteps === totalSteps,
    completedSteps: completedSteps,
    totalSteps: totalSteps,
    status: setupStatus,
    errors: setupErrors
  };
}

// ==========================================
// 段階的テスト機能
// ==========================================

/**
 * STEP1: 基本設定の確認
 * スクリプトプロパティの設定内容を検証し、現在の設定状況を表示
 * @return {boolean} 設定が正しい場合はtrue
 */
function step1_validateSetup() {
  logInfo('=== STEP1: 基本設定確認を開始します ===');
  
  // 設定の検証
  const configCheck = validateConfiguration();
  
  if (configCheck.errors.length > 0) {
    logError('step1_validateSetup', '設定にエラーがあります');
    configCheck.errors.forEach(error => Logger.log(`${error}`));
    return false;
  }
  
  // 設定値の表示（トークンは部分的にマスク）
  const maskedToken = mainAccessToken ? mainAccessToken.substring(0, 8) + '...' : 'null';
  Logger.log('');
  Logger.log('現在の設定:');
  Logger.log(`LINE_MESSAPI_TOKEN_MAIN: ${maskedToken}`);
  Logger.log(`LINE_GROUP_ID_MAIN: ${mainGroupId}`);
  Logger.log(`MAIL_GET_INTERVAL_MINUTE: ${emailFetchIntervalMinutes}分`);
  Logger.log(`Gmail検索ラベル: ${GMAIL_SEARCH_LABEL}`);
  
  logSuccess('STEP1: 基本設定確認が完了しました');
  Logger.log('');
  Logger.log('次のステップ: step2_testLineConnection() を実行してください');
  return true;
}

/**
 * STEP2: LINE Messaging API接続テスト
 * LINEアクセストークンの認証確認とテストメッセージの送信
 * @return {boolean} 接続テストが成功した場合はtrue
 */
function step2_testLineConnection() {
  logInfo('=== STEP2: LINE Messaging API接続テストを開始します ===');
  
  // 前提条件チェック
  const configCheck = validateConfiguration();
  if (!configCheck.isValid) {
    logError('step2_testLineConnection', '設定に問題があります。先にstep1_validateSetup()を実行してください');
    return false;
  }
  
  try {
    // 使用可能メッセージ数の確認
    logInfo('LINE Messaging API使用量を確認しています...');
    const remainingCount = getRemainingMessageCount(mainAccessToken);
    
    if (remainingCount <= 0) {
      logError('step2_testLineConnection', '1日の送信上限に達しているか、認証に失敗しました');
      return false;
    }
    
    logSuccess(`LINE Messaging API使用量確認完了: 残り${remainingCount}/${LINE_API_MAX_MESSAGES_PER_MONTH}通`);
    
    // テストメッセージ送信
    logInfo('テストメッセージを送信しています...');
    const testMessage = `【接続テスト】LINE Messaging API接続テストが成功しました 送信日時: ${new Date().toLocaleString('ja-JP')}`;
    
    const result = sendLineMessage(testMessage);
    
    if (result.success) {
      logSuccess('STEP2: LINE Messaging API接続テストが成功しました');
      Logger.log('');
      Logger.log('LINEグループでテストメッセージを受信できているか確認してください');
      Logger.log('次のステップ: step3_testGmailFetch() を実行してください');
      return true;
    } else {
      logError('step2_testLineConnection', `LINE Messaging API接続テストが失敗しました: ${result.errorMessage}`);
      return false;
    }
  } catch (error) {
    logError('step2_testLineConnection', 'テスト中にエラーが発生しました', error);
    return false;
  }
}

/**
 * STEP3: Gmail取得テスト
 * Gmailから対象ラベル付きメールの取得確認
 * @return {boolean} メール取得テストが成功した場合はtrue
 */
function step3_testGmailFetch() {
  logInfo('=== STEP3: Gmail取得テストを開始します ===');
  
  // 前提条件チェック
  const configCheck = validateConfiguration();
  if (!configCheck.isValid) {
    logError('step3_testGmailFetch', '設定に問題があります。先にstep1_validateSetup()を実行してください');
    return false;
  }
  
  try {
    logInfo('過去24時間のメールを検索しています...');
    const emails = fetchGmailMessages(60 * 24);
    
    if (emails.length === 0) {
      logInfo('過去24時間に対象メールは見つかりませんでした');
      return false;
    }
    
    logSuccess(`Gmail取得テスト完了: ${emails.length}件のメールが見つかりました`);
    
    // 取得したメールの詳細表示
    Logger.log('');
    Logger.log('取得されたメール:');
    emails.forEach((email, index) => {
      Logger.log(`--- メール ${index + 1} ---`);
      Logger.log(email);
      Logger.log('');
    });
    
    Logger.log('次のステップ: step4_testFullFlow() を実行してください');
    return true;
    
  } catch (error) {
    logError('step3_testGmailFetch', 'Gmail取得テスト中にエラーが発生しました', error);
    return false;
  }
}

/**
 * STEP4: 全体フローテスト
 * Gmail取得からLINE送信まで実際の運用フローをテスト実行
 * @return {boolean} 全体テストが成功した場合はtrue
 */
function step4_testFullFlow() {
  logInfo('=== STEP4: 全体フローテストを開始します ===');
  
  // 前提条件チェック
  const configCheck = validateConfiguration();
  if (!configCheck.isValid) {
    logError('step4_testFullFlow', '設定に問題があります。先にstep1_validateSetup()を実行してください');
    return false;
  }
  
  try {
    // Gmail取得テスト
    logInfo('テスト用メール検索を実行しています...');
    const emails = fetchGmailMessages(60 * 24); // 過去24時間
    
    if (emails.length === 0) {
      logInfo('テスト対象のメールが見つかりません');
      Logger.log('step3_testGmailFetch()を先に実行してメールが取得できることを確認してください');
      return false;
    }
    
    // 最初の1件のみをテスト送信
    const testEmail = emails[0];
    logInfo('テスト送信を実行しています...');
    
    const result = sendLineMessage(`【全体テスト】以下は実際のメール通知のサンプルです: ${testEmail}`);
    
    if (result.success) {
      logSuccess('STEP4: 全体フローテストが成功しました');
      Logger.log('');
      Logger.log('全ての設定が正しく動作しています');
      Logger.log('');
      Logger.log('本番運用の準備:');
      Logger.log('1. トリガーを設定して executeMain() 関数を定期実行');
      Logger.log(`2. 実行間隔を ${emailFetchIntervalMinutes}分 に設定`);
      Logger.log(`3. 通知したいメールに「${GMAIL_SEARCH_LABEL}」ラベルを付ける`);

      return true;
    } else {
      logError('step4_testFullFlow', `全体テストが失敗しました: ${result.errorMessage}`);
      return false;
    }
    
  } catch (error) {
    logError('step4_testFullFlow', '全体フローテスト中にエラーが発生しました', error);
    return false;
  }
}

// ==========================================
// メイン実行機能
// ==========================================

/**
 * メイン実行関数
 * トリガーから定期実行される主要な処理。Gmailからメールを取得してLINEに送信
 */
function executeMain() {
  logInfo('Gmail-LINE通知処理を開始します');
  
  try {
    // 設定の検証
    const configCheck = validateConfiguration();
    if (!configCheck.isValid) {
      logError('executeMain', '設定に問題があります');
      configCheck.errors.forEach(error => Logger.log(`設定エラー: ${error}`));
      return;
    }
    
    const emailMessages = fetchGmailMessages(emailFetchIntervalMinutes);

    if(emailMessages.length > 0){
      logInfo(`${emailMessages.length}件のメールを処理します`);
      let successCount = 0;
      let failureCount = 0;
      
      for(const message of emailMessages) {
        try {
          const result = sendLineMessage(message);
          if (result.success) {
            successCount++;
          } else {
            failureCount++;
          }
        } catch (error) {
          failureCount++;
          logError('executeMain', 'メッセージ送信中にエラーが発生しました', error);
        }
      }
      
      logInfo(`処理完了: 成功 ${successCount}件, 失敗 ${failureCount}件`);
    } else {
      logInfo('新しいメールはありませんでした');
    }
  } catch (error) {
    logError('executeMain', 'メイン処理中に予期しないエラーが発生しました', error);
  }
  
  logInfo('Gmail-LINE通知処理が完了しました');
}

/**
 * 動作確認用関数
 * 実際の送信は行わず、メール取得とフォーマットの確認のみ実行
 */
function executeOperationCheck() {
  logInfo('動作確認モードを開始します（過去24時間のメールを確認）');
  
  try {
    // 設定の検証
    const configCheck = validateConfiguration();
    if (!configCheck.isValid) {
      logError('executeOperationCheck', '設定に問題があります');
      configCheck.errors.forEach(error => Logger.log(`設定エラー: ${error}`));
      return;
    }
    
    const emailMessages = fetchGmailMessages(60*24*1);
    
    if(emailMessages.length > 0){
      logInfo(`${emailMessages.length}件のメールが見つかりました:`);
      for(let i = 0; i < emailMessages.length; i++) {
        Logger.log(`--- メール ${i+1} ---`);
        Logger.log(emailMessages[i]);
        Logger.log(''); // 空行
      }
    } else {
      logInfo('過去24時間に対象メールはありませんでした');
      Logger.log(`検索条件: ラベル「${GMAIL_SEARCH_LABEL}」が付いたメール`);
      Logger.log('メールにラベルが正しく付いているか確認してください');
    }
  } catch (error) {
    logError('executeOperationCheck', '動作確認処理中にエラーが発生しました', error);
  }
  
  logInfo('動作確認モードが完了しました');
}

/**
 * LINE送信テスト関数
 * 設定確認後にテストメッセージを送信
 */
function executeLineMessageTest() {
  logInfo('LINE送信テストを開始します');
  
  try {
    // 設定の検証
    const configCheck = validateConfiguration();
    if (!configCheck.isValid) {
      logError('executeLineMessageTest', '設定に問題があります');
      configCheck.errors.forEach(error => Logger.log(`設定エラー: ${error}`));
      return;
    }
    
    const testMessage = "【テスト】Gmail-LINE通知スクリプトのテストメッセージです " +
                       `送信日時: ${new Date().toLocaleString('ja-JP')}`;
    
    const result = sendLineMessage(testMessage);
    
    if (result.success) {
      logInfo('LINE送信テストが正常に完了しました');
    } else {
      logError('executeLineMessageTest', `LINE送信テストが失敗しました: ${result.errorMessage}`);
    }
  } catch (error) {
    logError('executeLineMessageTest', 'LINE送信テスト中にエラーが発生しました', error);
  }
}

// ==========================================
// LINE Messaging API 関連機能
// ==========================================

/**
 * LINE Messaging API残りメッセージ数取得
 * 1ヶ月あたりの送信制限に対する残り送信可能数を確認
 * @param {string} accessToken - LINE APIアクセストークン
 * @return {number} 残りメッセージ数（取得失敗時は0）
 */
function getRemainingMessageCount(accessToken) {
  if (!accessToken || accessToken.length === 0) {
    logError('getRemainingMessageCount', 'アクセストークンが無効です');
    return 0;
  }

  const options = {
    'method': 'get',
    'headers': createLineApiHeaders(accessToken)
  };

  try {
    const response = UrlFetchApp.fetch(LINE_API_QUOTA_URL, options);
    
    if (response.getResponseCode() !== 200) {
      logError('getRemainingMessageCount', `LINE API呼び出しが失敗しました。HTTPステータス: ${response.getResponseCode()}`);
      return 0;
    }
    
    const responseData = JSON.parse(response.getContentText());
    const totalUsage = responseData.totalUsage;

    if(Number.isInteger(totalUsage)) {
      const remaining = LINE_API_MAX_MESSAGES_PER_MONTH - totalUsage;
      logInfo(`LINE API残り送信数: ${remaining}/${LINE_API_MAX_MESSAGES_PER_MONTH}`);
      return remaining;
    } else {
      logError('getRemainingMessageCount', 'LINE APIのレスポンス形式が不正です');
      return 0;
    }
  } catch (error) {
    logError('getRemainingMessageCount', 'LINE API使用量取得中にエラーが発生しました', error);
    return 0;
  }
}

/**
 * LINEメッセージ送信実行
 * 複数のアクセストークンから使用可能なものを選択してメッセージを送信
 * @param {string} message - 送信するメッセージテキスト
 * @return {Object} 送信結果 {success: boolean, errorMessage: string}
 */
function sendLineMessage(message) {
  if (!message || message.trim().length === 0) {
    logError('sendLineMessage', '送信するメッセージが空です');
    return { success: false, errorMessage: 'メッセージが空です' };
  }

  // 送信可能なトークンを検索
  // 注意: 残り送信数の確認で LINE API を呼ぶため、頻繁に実行するとログや通信が増えます
  const availableTokenGroup = lineApiTokenGroups.find(elements => 
    elements.token && elements.group_id && 0 < getRemainingMessageCount(elements.token)
  );

  // 使用可能なトークンがない場合
  if(typeof availableTokenGroup === "undefined") {
    const errorMsg = '送信可能なLINE APIトークンがありません（1日の送信上限に達したか、設定が無効の可能性があります）';
    logError('sendLineMessage', errorMsg);
    return { success: false, errorMessage: errorMsg };
  }

  const headers = createLineApiHeaders(availableTokenGroup.token);
  const data = {
    'to': availableTokenGroup.group_id,
    'messages': [
      {
        'type': 'text',
        'text': message
      }
    ]
  };

  const options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(data)
  };

  try {
    const response = UrlFetchApp.fetch(LINE_API_PUSH_URL, options);
    
    if (response.getResponseCode() === 200) {
      logInfo('LINEメッセージ送信が完了しました');
      return { success: true, errorMessage: null };
    } else {
      const errorMsg = `LINE API呼び出しが失敗しました。HTTPステータス: ${response.getResponseCode()}`;
      logError('sendLineMessage', errorMsg);
      logError('sendLineMessage', `レスポンス内容: ${response.getContentText()}`);
      return { success: false, errorMessage: errorMsg };
    }
  } catch (error) {
    logError('sendLineMessage', 'LINEメッセージ送信中にエラーが発生しました', error);
    return { success: false, errorMessage: error.toString() };
  }
}

// ==========================================
// Gmail関連機能
// ==========================================

/**
 * Gmailメッセージ取得処理
 * 指定した時間範囲内の対象ラベル付きメールを取得し、整形して返す
 * @param {number} intervalMinutes - 取得対象期間（分）
 * @return {Array} 整形済みメッセージ配列
 */
function fetchGmailMessages(intervalMinutes) {
  if (!intervalMinutes || isNaN(intervalMinutes) || intervalMinutes <= 0) {
    logError('fetchGmailMessages', '取得間隔が無効です');
    return [];
  }

  try {
    // 取得間隔を元に取得するメールの期間を計算
    const now = Math.floor(new Date().getTime() / 1000);
    const intervalSeconds = now - (60 * intervalMinutes);

    // Gmail検索条件を作成（指定されたラベルが付いたメールを検索）
    const searchTerms = `(after:${intervalSeconds} label:${GMAIL_SEARCH_LABEL})`;
    logInfo(`Gmail検索条件: ${searchTerms}`);

    // Gmailからメール取得
    const threads = GmailApp.search(searchTerms);
    const emails = GmailApp.getMessagesForThreads(threads);

    logInfo(`取得したメールスレッド数: ${emails.length}件`);

    const formattedMessages = [];

    if(emails && emails.length > 0) {
      for (let i = 0; i < emails.length; i++) {
        try {
          const emailThread = emails[i];
          
          if (!emailThread || emailThread.length === 0) {
            logError('fetchGmailMessages', `メールスレッド ${i+1} が空です`);
            continue;
          }
          
          // 各スレッドの最新メールを取得
          const latestEmailIndex = emailThread.length - 1;
          const email = emailThread[latestEmailIndex];
          
          // 日時情報を整形
          const emailDate = email.getDate();
          const formattedDateTime = `${emailDate.getMonth() + 1}月${emailDate.getDate()}日 ${emailDate.getHours()}時${("00" + emailDate.getMinutes()).slice(-2)}分`;

          // メール件名と本文を整形
          const formattedSubject = formatEmailSubject(email.getSubject());
          const formattedBody = formatEmailBody(email.getPlainBody());

          // 最終メッセージを組み立て
          let finalMessage = formattedDateTime;
          if(formattedSubject && formattedSubject.trim().length > 0) {
            finalMessage += " 「" + formattedSubject + "」";
          }
          finalMessage += " " + formattedBody;

          formattedMessages.push(finalMessage);
          
        } catch (error) {
          logError('fetchGmailMessages', `メール ${i+1} の処理中にエラーが発生しました`, error);
        }
      }
    }

    return formattedMessages;
  } catch(error) {
    logError('fetchGmailMessages', 'Gmailメッセージ取得中にエラーが発生しました', error);
    return [];
  }
}

// ==========================================
// テキスト処理関数
// ==========================================

/**
 * メール件名フォーマット処理
 * 設定されたパターンに基づいて件名から不要な文字列を除去
 * @param {string} subject - 元の件名
 * @return {string} 整形済み件名
 */
function formatEmailSubject(subject) {
  if (!subject) {
    return "";
  }

  try {
    let cleanedText = String(subject);

    if(cleanedText.length > 0) {
      // 指定パターンを削除
      SUBJECT_DELETE_PATTERNS.forEach(pattern => {
        if (!pattern || !pattern.trim()) return;
        cleanedText = removeSpecificText(cleanedText, pattern);
      });

      // 指定パターン以降を削除
      SUBJECT_DELETE_TO_END_PATTERNS.forEach(pattern => {
        if (!pattern || !pattern.trim()) return;
        cleanedText = removeTextFromPattern(cleanedText, pattern);
      });
    }

    return cleanedText.trim();
  } catch (error) {
    logError('formatEmailSubject', '件名の整形中にエラーが発生しました', error);
    return String(subject); // エラー時は元の件名を返す
  }
}

/**
 * メール本文フォーマット処理
 * 設定されたパターンに基づいて本文から不要な文字列を除去し、フォーマットを整える
 * @param {string} body - 元の本文
 * @return {string} 整形済み本文
 */
function formatEmailBody(body) {
  if (!body) {
    return "";
  }

  try {
    let cleanedText = String(body);

    if(cleanedText.length > 0) {
      // 指定パターンを削除
      BODY_DELETE_PATTERNS.forEach(pattern => {
        if (!pattern || !pattern.trim()) return;
        cleanedText = removeSpecificText(cleanedText, pattern);
      });

      // 指定パターン以降を削除
      BODY_DELETE_TO_END_PATTERNS.forEach(pattern => {
        if (!pattern || !pattern.trim()) return;
        cleanedText = removeTextFromPattern(cleanedText, pattern);
      });
      
      // 空白文字を削除
      cleanedText = removeWhitespaceCharacters(cleanedText);

      // 時間表記を短縮（特定システム用の処理）
      cleanedText = shortenTimeFormat(cleanedText);

      // 連続改行を正規化
      cleanedText = normalizeConsecutiveNewlines(cleanedText);
    }

    return cleanedText.trim();
  } catch (error) {
    logError('formatEmailBody', '本文の整形中にエラーが発生しました', error);
    return String(body); // エラー時は元の本文を返す
  }
}

/**
 * 指定文字列以降削除処理
 * 指定された文字列が見つかった位置から文末まですべて削除
 * @param {string} message - 処理対象の文字列
 * @param {string} searchPattern - 削除開始位置の文字列
 * @return {string} 処理後の文字列
 */
function removeTextFromPattern(message, searchPattern) {
  if (!message || !searchPattern) {
    return message || "";
  }

  try {
    // パターンの位置を検索
    const index = message.indexOf(searchPattern);
    // パターンが見つかった場合、その位置までの文字列を取得
    const result = index !== -1 ? message.substring(0, index) : message;
    // 前後の空白を削除して返す
    return result.trim();
  } catch (error) {
    logError('removeTextFromPattern', '文字列削除処理中にエラーが発生しました', error);
    return message; // エラー時は元の文字列を返す
  }
}

/**
 * 特定文字列削除処理
 * 指定された文字列をすべて除去
 * @param {string} message - 処理対象の文字列
 * @param {string} searchPattern - 削除する文字列
 * @return {string} 処理後の文字列
 */
function removeSpecificText(message, searchPattern) {
  if (!message || !searchPattern) return message || "";

  try {
    // 全ての出現箇所を削除（正規表現ではなく文字列一致）
    return message.split(searchPattern).join("");
  } catch (error) {
    logError('removeSpecificText', '特定文字列削除処理中にエラーが発生しました', error);
    return message;
  }
}

/**
 * 連続改行正規化処理
 * 複数の連続した改行文字を単一の改行に統一
 * @param {string} input - 処理対象の文字列
 * @return {string} 処理後の文字列
 */
function normalizeConsecutiveNewlines(input) {
  if (!input) {
    return "";
  }

  try {
    // \r\nの改行を\nに置換してから実行
    // 2個以上の改行を正規表現で検出し、1つの改行に置き換える
    return input.replace(/\r\n/g, '\n').replace(/\n{2,}/g, '\n');
  } catch (error) {
    logError('normalizeConsecutiveNewlines', '改行正規化処理中にエラーが発生しました', error);
    return input; // エラー時は元の文字列を返す
  }
}

/**
 * 時間フォーマット短縮処理
 * 特定の時間表記フォーマット（年月日時分秒）を時分のみに短縮
 * @param {string} message - 処理対象の文字列
 * @return {string} 時と分に短縮した文字列
 */
function shortenTimeFormat(message) {
  if (!message) {
    return "";
  }

  try {
    const timePattern = /(\d{4}年\d{2}月\d{2}日)(\d{1,2})時(\d{1,2})分\d{1,2}秒/;
    return message.replace(timePattern, (match, p1, p2, p3) => {
      const hours = parseInt(p2, 10);
      const minutes = parseInt(p3, 10);
      return `${hours}時${minutes}分`;
    });
  } catch (error) {
    logError('shortenTimeFormat', '時間フォーマット変換処理中にエラーが発生しました', error);
    return message; // エラー時は元の文字列を返す
  }
}

/**
 * 空白文字除去処理
 * 全角スペース、半角スペース、タブ文字を除去し、前後の空白をトリム
 * @param {string} str - 処理対象の文字列
 * @return {string} 処理後の文字列
 */
function removeWhitespaceCharacters(str) {
  if (!str) {
    return "";
  }

  try {
    return str.replace(/[　 \t]/g, "").trim();
  } catch (error) {
    logError('removeWhitespaceCharacters', '空白文字削除処理中にエラーが発生しました', error);
    return str; // エラー時は元の文字列を返す
  }
}
