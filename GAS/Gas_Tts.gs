/**
 * Google Apps Script for Cloud Text-to-Speech API
 * Everyday English アプリ用 - TTS機能
 * 
 * 注意: エントリーポイント（doPost）は Gas_Main.gs に統合されています。
 * このファイルにはTTS関連の関数のみが含まれます。
 */

/**
 * テキストから言語を自動判定
 * 日本語文字（ひらがな、カタカナ、漢字）が含まれていれば日本語、そうでなければ英語
 */
function detectLanguage(text) {
  // 日本語文字の正規表現（ひらがな、カタカナ、漢字）
  var japanesePattern = /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/;
  
  if (japanesePattern.test(text)) {
    return 'ja-JP';
  } else {
    return 'en-US';
  }
}

/**
 * 言語コードから音声名を取得
 */
/**
 * 言語コードと性別に応じて音声名を取得
 * @param {string} languageCode - 言語コード（'ja-JP' または 'en-US'）
 * @param {string} voiceGender - 音声の性別（'male' または 'female'、デフォルト：'female'）
 * @return {string} 音声名
 */
function getVoiceName(languageCode, voiceGender) {
  // voiceGenderが未定義やnullの場合は'female'として扱う
  if (!voiceGender || typeof voiceGender !== 'string') {
    voiceGender = 'female';
  }
  
  if (languageCode === 'ja-JP') {
    if (voiceGender === 'male') {
      return 'ja-JP-Neural2-D'; // 日本語男性
    } else {
      return 'ja-JP-Neural2-B'; // 日本語女性
    }
  } else {
    if (voiceGender === 'male') {
      return 'en-US-Neural2-J'; // 英語男性
    } else {
      return 'en-US-Neural2-F'; // 英語女性
    }
  }
}

/**
 * TTS用にテキストを正規化
 * - 前後の空白を削除
 * - 特殊な空白文字を通常のスペースに変換
 * - 連続する空白を1つに正規化
 */
function normalizeTextForTTS(text) {
  if (!text || typeof text !== 'string') {
    return '';
  }
  
  // 1. 前後の空白を削除
  var normalized = text.trim();
  
  // 2. 特殊な空白文字を通常のスペース（U+0020）に変換
  // 全角スペース（U+3000）、タブ（U+0009）、改行（U+000A, U+000D）、
  // ノンブレーキングスペース（U+00A0）などを通常のスペースに変換
  normalized = normalized.replace(/[\u3000\u0009\u000A\u000D\u00A0\u2000-\u200B\u2028\u2029]/g, ' ');
  
  // 3. 連続する空白を1つに正規化（2文字以上の空白を1文字に）
  normalized = normalized.replace(/\s+/g, ' ');
  
  return normalized;
}

/**
 * テキストをSSML形式に変換
 * - 空白を <break> タグに置き換えて、読み上げを防止
 * - SSMLの特殊文字をエスケープ
 */
function convertTextToSSML(text) {
  if (!text || typeof text !== 'string') {
    return '';
  }
  
  // 1. SSMLの特殊文字をエスケープ（& を最初に処理する必要がある）
  var ssml = text
    .replace(/&/g, '&amp;')  // & を &amp; に変換（最初に処理）
    .replace(/</g, '&lt;')   // < を &lt; に変換
    .replace(/>/g, '&gt;');  // > を &gt; に変換
  
  // 2. 空白を <break time="0.0s"/> に置き換え
  // 0.0秒のポーズを挿入して、空白が読み上げられるのを防ぐ（効果検証用）
  ssml = ssml.replace(/\s/g, '<break time="0.0s"/>');
  
  // 3. SSMLタグでラップ
  ssml = '<speak>' + ssml + '</speak>';

  return ssml;
}

/**
 * Cloud Text-to-Speech APIを呼び出す
 * @param {string} text - 読み上げるテキスト
 * @param {string} languageCode - 言語コード（'ja-JP' または 'en-US'）
 * @param {string} voiceName - 音声名
 * @param {string} voiceGender - 音声の性別（'male' または 'female'、デフォルト：'female'）
 * @param {string} speed - 読み上げの速さ（'fast', 'medium', 'slow'、デフォルト：'fast'）
 */
function callTextToSpeechAPI(text, languageCode, voiceName, voiceGender, speed) {
  // テキストをSSML形式に変換（空白を <break> タグに置き換え）
  var ssml = convertTextToSSML(text);
  
  // APIキーをPropertiesServiceから取得
  var apiKey = PropertiesService.getScriptProperties().getProperty('TTS_API_KEY');
  
  if (!apiKey) {
    throw new Error('API key is not set. Please run setApiKey() function first.');
  }
  
  // 音声の性別を設定（デフォルト値：'female'）
  // voiceGenderが未定義やnullの場合は'female'として扱う
  if (!voiceGender || typeof voiceGender !== 'string') {
    voiceGender = 'female';
  }
  var gender = (voiceGender === 'male') ? 'MALE' : 'FEMALE';
  
  // 読み上げの速さを設定（デフォルト値：1.25）
  // speedが未定義やnullの場合は'fast'として扱う
  if (!speed || typeof speed !== 'string') {
    speed = 'fast';
  }
  var speedMap = {
    'fast': 1.25,
    'medium': 1.20,
    'slow': 1.15
  };
  var speakingRate = speedMap[speed];
  if (speakingRate === undefined) {
    speakingRate = 1.25; // デフォルト値
  }
  
  // Cloud Text-to-Speech APIのエンドポイント
  var url = 'https://texttospeech.googleapis.com/v1/text:synthesize?key=' + apiKey;
  
  // リクエストボディ（SSML形式を使用）
  // 注意: nameを指定する場合はssmlGenderは不要（音声名に性別が含まれている）
  var requestBody = {
    input: {
      ssml: ssml  // text の代わりに ssml を使用
    },
    voice: {
      languageCode: languageCode,
      name: voiceName
      // ssmlGenderは使用しない（nameに性別が含まれているため）
    },
    audioConfig: {
      audioEncoding: 'MP3',
      speakingRate: speakingRate,
      pitch: 0.0,
      volumeGainDb: 0.0
    }
  };
  
  // APIを呼び出し
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  
  if (responseCode !== 200) {
    // 詳細なエラー情報はログに記録（内部のみ）
    Logger.log('[Error] TTS API error - Code: ' + responseCode + ', Response: ' + responseText);
    // ユーザーには一般的なエラーメッセージを返す
    throw new Error('TTS API request failed');
  }
  
  var result = JSON.parse(responseText);
  
  if (!result.audioContent) {
    throw new Error('No audio content in response');
  }
  
  return result.audioContent;
}


/**
 * APIキーを設定する方法:
 * 
 * 【推奨方法】Google Apps Scriptエディタから直接設定:
 * 1. メニューから「プロジェクトの設定」を開く
 * 2. 「スクリプト プロパティ」セクションを開く
 * 3. 「スクリプト プロパティを追加」をクリック
 * 4. プロパティ名: TTS_API_KEY
 * 5. プロパティ値: 実際のCloud Text-to-Speech APIのAPIキー
 * 6. 「保存」をクリック
 * 
 * この方法なら、APIキーをコード内に記述する必要がありません。
 */

