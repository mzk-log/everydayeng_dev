/**
 * Google Apps Script - Main Entry Point
 * Everyday English アプリ用 - 統合エントリーポイント
 * 
 * TTS（音声生成）とDATA（データ取得）の両方を処理する統合エントリーポイント
 */

// 許可されたドメインのリスト
var ALLOWED_DOMAINS = [
  'https://mzk-log.github.io',
  'http://localhost',
  'http://127.0.0.1',
  'file://' // ローカルファイル実行時用
];

// テキストの最大長さ（文字数）- TTS用
var MAX_TEXT_LENGTH = 1000;

/**
 * Webアプリとして公開するメイン関数（GETリクエスト）
 * DATA処理のみをサポート（TTSはPOSTのみ）
 */
function doGet(e) {
  try {
    // 1. リファラーチェック
    var refererError = validateReferer(e.parameter.referer || '');
    if (refererError) {
      return refererError;
    }
    
    // 2. メールアドレス認証
    var emailError = validateEmail(e.parameter.email || '');
    if (emailError) {
      return emailError;
    }
    
    // 3. アクションに応じて処理を分岐
    var action = e.parameter.action || '';
    
    if (action === 'getCategories') {
      return getCategories();
    } else if (action === 'getCategoryData') {
      var categoryNo = e.parameter.categoryNo || '';
      if (!categoryNo) {
        return createErrorResponse('categoryNo parameter is required');
      }
      return getCategoryData(categoryNo);
    } else {
      // GETリクエストでは更新系のactionは許可しない（セキュリティのため）
      return createErrorResponse('Invalid action parameter');
    }
    
  } catch (error) {
    // 詳細なエラー情報はログに記録（内部のみ）
    Logger.log('[Error] doGet exception: ' + error.toString());
    Logger.log('[Error] Stack trace: ' + error.stack);
    // ユーザーには一般的なエラーメッセージを返す
    return createErrorResponse('An error occurred while processing your request');
  }
}

/**
 * Webアプリとして公開するメイン関数（POSTリクエスト）
 * TTS（音声生成）とDATA（データ取得）の両方を処理
 */
function doPost(e) {
  try {
    // POSTリクエストのパラメータを取得
    // e.parameter（URLパラメータ）と e.postData.contents（POSTボディ）の両方をチェック
    var params = {};
    
    // URLパラメータから取得
    if (e.parameter) {
      for (var key in e.parameter) {
        params[key] = e.parameter[key];
      }
    }
    
    // POSTボディから取得（application/x-www-form-urlencoded形式）
    if (e.postData && e.postData.contents) {
      var postParams = e.postData.contents.split('&');
      for (var i = 0; i < postParams.length; i++) {
        var pair = postParams[i].split('=');
        if (pair.length === 2) {
          var key = decodeURIComponent(pair[0]);
          var value = decodeURIComponent(pair[1]);
          params[key] = value;
        }
      }
    }
    
    // 1. リファラーチェック
    var refererError = validateReferer(params.referer || '');
    if (refererError) {
      return refererError;
    }
    
    // 2. パラメータに応じて処理を分岐
    var text = params.text || '';
    var action = params.action || '';
    
    // + を空白に置き換え（URLSearchParamsのエンコード形式に対応）
    // application/x-www-form-urlencoded形式では、空白は+にエンコードされるが、
    // decodeURIComponent()は+を空白にデコードしないため、手動で置き換える必要がある
    if (text) {
      text = text.replace(/\+/g, ' ');
    }
    
    // textパラメータがある場合 → TTS処理（音声生成）
    if (text && text.trim() !== '') {
      var emailError = validateEmail(params.email || '', 'TTS');
      if (emailError) {
        return emailError;
      }
      return processTTS(text, params);
    }
    
    // actionパラメータがある場合 → DATA処理（データ取得・更新）
    if (action) {
      var emailError = validateEmail(params.email || '');
      if (emailError) {
        return emailError;
      }
      
      if (action === 'getCategories') {
        return getCategories();
      } else if (action === 'getCategoryData') {
        var categoryNo = params.categoryNo || '';
        if (!categoryNo) {
          return createErrorResponse('categoryNo parameter is required');
        }
        return getCategoryData(categoryNo);
      } else if (action === 'updateAnswerMemo') {
        var id = params.id || '';
        var answer = params.answer || '';
        if (!id) {
          return createErrorResponse('id parameter is required');
        }
        return updateAnswerMemo(id, answer);
      } else if (action === 'speechToText') {
        var audioContent = params.audioContent || '';
        var languageCode = params.languageCode || 'ja-JP';
        if (!audioContent) {
          return createErrorResponse('audioContent parameter is required');
        }
        return processSpeechToTextRequest(audioContent, languageCode);
      } else {
        return createErrorResponse('Invalid action parameter');
      }
    }
    
    // どちらのパラメータもない場合
    return createErrorResponse('Invalid request: text or action parameter is required');
    
  } catch (error) {
    // 詳細なエラー情報はログに記録（内部のみ）
    Logger.log('[Error] doPost exception: ' + error.toString());
    Logger.log('[Error] Stack trace: ' + error.stack);
    // ユーザーには一般的なエラーメッセージを返す
    return createErrorResponse('An error occurred while processing your request');
  }
}

/**
 * TTS処理（音声生成）
 * @param {string} text - 読み上げるテキスト
 * @param {Object} params - リクエストパラメータ（voiceGender, speedを含む）
 */
function processTTS(text, params) {
  try {
    // テキストを正規化（空白の処理）
    text = normalizeTextForTTS(text);
    
    // 入力検証
    if (text.length > MAX_TEXT_LENGTH) {
      return createErrorResponse('Text length exceeds maximum limit of ' + MAX_TEXT_LENGTH + ' characters');
    }
    
    // 言語を自動判定
    var languageCode = detectLanguage(text);
    
    // 音声設定を取得（デフォルト値：女性、fast）
    var voiceGender = (params && params.voiceGender) ? params.voiceGender : 'female';
    var speed = (params && params.speed) ? params.speed : 'fast';
    
    // 性別に応じて音声名を取得
    var voiceName = getVoiceName(languageCode, voiceGender);
    
    // Cloud Text-to-Speech APIを呼び出し
    var audioContent = callTextToSpeechAPI(text, languageCode, voiceName, voiceGender, speed);
    
    if (!audioContent) {
      return createErrorResponse('Failed to generate speech');
    }
    
    // 音声データを返す
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      audioContent: audioContent,
      language: languageCode,
      voiceName: voiceName
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // 詳細なエラー情報はログに記録（内部のみ）
    Logger.log('[Error] processTTS exception: ' + error.toString());
    Logger.log('[Error] Stack trace: ' + error.stack);
    // ユーザーには一般的なエラーメッセージを返す
    return createErrorResponse('Failed to generate speech: ' + error.toString());
  }
}

/**
 * リファラーチェックを実行
 * @param {string} referer - リファラー文字列
 * @return {Object|null} - エラーの場合はエラーレスポンス、正常な場合はnull
 */
function validateReferer(referer) {
  var decodedReferer = referer || '';
  
  // URLデコード処理
  try {
    decodedReferer = decodeURIComponent(decodedReferer);
  } catch (decodeError) {
    // デコードに失敗した場合は元の値を使用
  }
  
  if (!isAllowedDomain(decodedReferer)) {
    return createErrorResponse('Access denied: Invalid referer');
  }
  
  return null; // 正常な場合はnullを返す
}

/**
 * メール認証を実行
 * @param {string} email - メールアドレス
 * @param {string} context - エラーメッセージのコンテキスト（'TTS' または 'DATA'、省略可）
 * @return {Object|null} - エラーの場合はエラーレスポンス、正常な場合はnull
 */
function validateEmail(email, context) {
  if (!email) {
    var errorMessage = 'Access denied: Email parameter is required';
    if (context === 'TTS') {
      errorMessage += ' for TTS';
    }
    return createErrorResponse(errorMessage);
  }
  
  if (!isAllowedEmail(email)) {
    return createErrorResponse('Access denied: Email not authorized. Please check if your email is registered in Script Properties (ALLOWED_EMAILS).');
  }
  
  return null; // 正常な場合はnullを返す
}

/**
 * 許可されたドメインかチェック
 * 【セキュリティ修正】
 * - 空のリファラーを拒否
 * - 完全一致チェック（部分一致ではなく、正確なドメインのみ許可）
 * - プロトコルを含めた完全一致チェック
 * - 環境ごとの制御（DEV:開発環境 / PRD:本番環境）
 */
function isAllowedDomain(referer) {
  // リファラーが空の場合は拒否（セキュリティ強化）
  if (!referer || referer.trim() === '') {
    return false;
  }
  
  var trimmedReferer = referer.trim();
  
  // 環境設定を取得（DEV:開発環境 / PRD:本番環境）
  // 設定方法: Script Propertiesに「ALLOW_FILE_PROTOCOL」を追加し、値に「DEV」または「PRD」を設定
  // デフォルト: PRD（セキュリティのため、設定がない場合は本番環境として扱う）
  var environment = PropertiesService.getScriptProperties().getProperty('ALLOW_FILE_PROTOCOL');
  var isDevEnvironment = environment && environment.toUpperCase() === 'DEV';
  
  // 開発環境専用の許可チェック（file://, localhost, 127.0.0.1）
  if (isDevEnvironment) {
    // file://プロトコルのチェック
    if (trimmedReferer.indexOf('file://') === 0) {
      return true; // file://で始まるすべてのURLを許可
    }
    
    // localhostのチェック（ポート番号を無視）
    if (trimmedReferer.indexOf('http://localhost') === 0 || trimmedReferer.indexOf('https://localhost') === 0) {
      return true;
    }
    
    // 127.0.0.1のチェック（ポート番号を無視）
    if (trimmedReferer.indexOf('http://127.0.0.1') === 0 || trimmedReferer.indexOf('https://127.0.0.1') === 0) {
      return true;
    }
  }
  
  // リファラーを正規化（前後の空白を削除、末尾のスラッシュを統一）
  var normalizedReferer = trimmedReferer;
  if (normalizedReferer.endsWith('/')) {
    normalizedReferer = normalizedReferer.slice(0, -1);
  }
  
  // URLを解析してプロトコルとホスト名を抽出（ポート番号を無視）
  try {
    // URL形式の場合（http://, https://）
    if (normalizedReferer.indexOf('http://') === 0 || normalizedReferer.indexOf('https://') === 0) {
      // プロトコルとホスト名を抽出（ポート番号を除去）
      var urlMatch = normalizedReferer.match(/^(https?:\/\/[^\/:]+)/);
      if (urlMatch) {
        var protocolAndHost = urlMatch[1];
        
        // 許可されたドメインのリストをチェック
        for (var i = 0; i < ALLOWED_DOMAINS.length; i++) {
          var allowedDomain = ALLOWED_DOMAINS[i].trim();
          
          // file://は開発環境でのみ許可（既に上でチェック済み）
          if (allowedDomain.indexOf('file://') === 0) {
            continue;
          }
          
          // localhostと127.0.0.1は開発環境でのみ許可（既に上でチェック済み）
          if (allowedDomain.indexOf('localhost') !== -1 || allowedDomain.indexOf('127.0.0.1') !== -1) {
            continue;
          }
          
          // 許可ドメインからもプロトコルとホスト名を抽出
          var allowedMatch = allowedDomain.match(/^(https?:\/\/[^\/:]+)/);
          if (allowedMatch) {
            var allowedProtocolAndHost = allowedMatch[1];
            
            // プロトコルとホスト名が一致するかチェック
            if (protocolAndHost === allowedProtocolAndHost) {
              return true;
            }
          }
        }
      }
    }
  } catch (e) {
    // URL解析に失敗した場合は従来の方法でチェック
  }
  
  // 従来のチェック（完全一致とサブパス）- フォールバック
  for (var j = 0; j < ALLOWED_DOMAINS.length; j++) {
    var allowedDomainFallback = ALLOWED_DOMAINS[j].trim();
    // 末尾のスラッシュを統一
    if (allowedDomainFallback.endsWith('/')) {
      allowedDomainFallback = allowedDomainFallback.slice(0, -1);
    }
    
    // file://, localhost, 127.0.0.1は開発環境でのみ許可（既に上でチェック済み）
    if (allowedDomainFallback.indexOf('file://') === 0 || 
        allowedDomainFallback.indexOf('localhost') !== -1 || 
        allowedDomainFallback.indexOf('127.0.0.1') !== -1) {
      continue;
    }
    
    // 完全一致チェック（プロトコルを含む）
    if (normalizedReferer === allowedDomainFallback) {
      return true;
    }
    
    // サブパスも許可（例: https://mzk-log.github.io/path/to/page）
    if (normalizedReferer.indexOf(allowedDomainFallback + '/') === 0) {
      return true;
    }
  }
  
  // 許可リストにない場合は false
  return false;
}

/**
 * エラーレスポンスを作成
 */
function createErrorResponse(message) {
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    error: message
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Speech-to-Text APIリクエストを処理
 * @param {string} audioContent - Base64エンコードされた音声データ
 * @param {string} languageCode - 言語コード
 */
function processSpeechToTextRequest(audioContent, languageCode) {
  try {
    var recognizedText = processSpeechToText(audioContent, languageCode);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      text: recognizedText
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // 詳細なエラー情報はログに記録（内部のみ）
    Logger.log('[Error] processSpeechToTextRequest exception: ' + error.toString());
    Logger.log('[Error] Stack trace: ' + error.stack);
    // ユーザーには一般的なエラーメッセージを返す
    return createErrorResponse('Failed to recognize speech: ' + error.toString());
  }
}


