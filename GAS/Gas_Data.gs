/**
 * Google Apps Script for Data Access (with Email Authentication)
 * Everyday English アプリ用 - データ取得機能
 * 
 * 注意: エントリーポイント（doGet, doPost）は Gas_Main.gs に統合されています。
 * このファイルにはデータ取得関連の関数のみが含まれます。
 */

// スプレッドシートID
var SPREADSHEET_ID = '1pnlKMrp07Yz4MMCFByw8F04ttT3Cf6xDSX7zf5R64ZA';

// キャッシュの有効期限（秒）
var CACHE_EXPIRY_SECONDS = 300; // 5分

  /**
  * 許可されたメールアドレスかチェック
  * 
  * メールアドレスの管理方法（Script Propertiesで管理）:
  * 1. Google Apps Scriptエディタでメニューから「プロジェクトの設定」を開く
  * 2. 「スクリプト プロパティ」セクションを開く
  * 3. 「スクリプト プロパティを追加」をクリック
  * 4. プロパティ名: ALLOWED_EMAILS
  * 5. プロパティ値: カンマ区切りでメールアドレスを設定（例: user1@example.com,user2@example.com）
  * 6. 「保存」をクリック
  * 
  * 注意: メールアドレスは大文字小文字を区別しません。前後の空白は自動的に削除されます。
  */
  function isAllowedEmail(email) {
    if (!email || email.trim() === '') {
      return false;
    }
    
    // メールアドレスの正規化（小文字に変換、前後の空白を削除）
    var normalizedEmail = email.toLowerCase().trim();
    
    // Script Propertiesから許可されたメールアドレスリストを取得
    var allowedEmailsStr = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS');
    
    if (!allowedEmailsStr || allowedEmailsStr.trim() === '') {
      return false;
    }
    
    // カンマ区切りで分割して、各メールアドレスを正規化
    var allowedEmails = allowedEmailsStr.split(',').map(function(e) {
      return e.trim().toLowerCase();
    });
    
    // メールアドレスが許可リストに含まれているかチェック
    for (var i = 0; i < allowedEmails.length; i++) {
      if (allowedEmails[i] === normalizedEmail) {
        return true;
      }
    }
    
    return false;
  }

  /**
  * カテゴリ一覧を取得
  */
  function getCategories() {
    try {
      // キャッシュをチェック
      var cache = CacheService.getScriptCache();
      var cacheKey = 'categories';
      var cachedData = cache.get(cacheKey);
      
      if (cachedData) {
        return ContentService.createTextOutput(cachedData).setMimeType(ContentService.MimeType.JSON);
      }
      
      // スプレッドシートからデータを取得
      var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      var dataSheet = sheet.getSheetByName('Data');
      
      if (!dataSheet) {
        return createErrorResponse('Data sheet not found');
      }
      
      // B列（Category_No）とC列（Category）を取得（2行目以降）
      var lastRow = dataSheet.getLastRow();
      if (lastRow < 2) {
        return createErrorResponse('No data found');
      }
      
      var range = dataSheet.getRange(2, 2, lastRow - 1, 2); // B2からC列の最後まで
      var values = range.getValues();
      
      // カテゴリを重複排除し、設問数をカウント
      var categoryMap = {};
      var categoriesList = [];
      
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        if (!row || row.length < 2) continue;
        
        var categoryNo = row[0]; // B列: Category_No
        var category = row[1];   // C列: Category
        
        // 空の場合はスキップ
        if (!categoryNo || !category) continue;
        
        // 数値の場合は文字列に変換
        if (typeof categoryNo === 'number') {
          categoryNo = String(categoryNo);
        }
        
        // 重複チェック
        if (!categoryMap[categoryNo]) {
          categoryMap[categoryNo] = {
            name: category,
            count: 0
          };
        }
        // 設問数をカウント
        categoryMap[categoryNo].count++;
      }
      
      // カテゴリリストを作成
      for (var categoryNo in categoryMap) {
        if (categoryMap.hasOwnProperty(categoryNo)) {
          categoriesList.push({
            no: categoryNo,
            name: categoryMap[categoryNo].name,
            count: categoryMap[categoryNo].count
          });
        }
      }
      
      if (categoriesList.length === 0) {
        return createErrorResponse('No categories found');
      }
      
      var result = {
        success: true,
        categories: categoriesList
      };
      
      var jsonResult = JSON.stringify(result);
      
      // キャッシュに保存
      cache.put(cacheKey, jsonResult, CACHE_EXPIRY_SECONDS);
      
      return ContentService.createTextOutput(jsonResult).setMimeType(ContentService.MimeType.JSON);
      
    } catch (error) {
      // 詳細なエラー情報はログに記録（内部のみ）
      Logger.log('[Error] getCategories exception: ' + error.toString());
      Logger.log('[Error] Stack trace: ' + error.stack);
      // ユーザーには一般的なエラーメッセージを返す
      return createErrorResponse('Failed to retrieve categories');
    }
  }

  /**
  * カテゴリデータを取得
  */
  function getCategoryData(categoryNo) {
    try {
      // カテゴリ番号を文字列に統一
      if (typeof categoryNo === 'number') {
        categoryNo = String(categoryNo);
      }
      
      // キャッシュをチェック
      var cache = CacheService.getScriptCache();
      var cacheKey = 'categoryData_' + categoryNo;
      var cachedData = cache.get(cacheKey);
      
      if (cachedData) {
        return ContentService.createTextOutput(cachedData).setMimeType(ContentService.MimeType.JSON);
      }
      
      // スプレッドシートからデータを取得
      var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      var dataSheet = sheet.getSheetByName('Data');
      
      if (!dataSheet) {
        return createErrorResponse('Data sheet not found');
      }
      
      // A列（ID）からI列（note）まで取得（2行目以降）
      var lastRow = dataSheet.getLastRow();
      if (lastRow < 2) {
        return createErrorResponse('No data found');
      }
      
      var range = dataSheet.getRange(2, 1, lastRow - 1, 9); // A2からI列の最後まで
      var values = range.getValues();
      
      var items = [];
      
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        if (!row || row.length < 2) continue;
        
        var rowCategoryNo = row[1]; // B列: Category_No
        
        // 数値の場合は文字列に変換
        if (typeof rowCategoryNo === 'number') {
          rowCategoryNo = String(rowCategoryNo);
        }
        
        // カテゴリ番号が一致する行のみ取得
        if (rowCategoryNo == categoryNo) {
          items.push({
            id: row[0] || '',           // A列: ID
            no: row[3] || '',           // D列: No
            q_title: row[4] || '',      // E列: Q_Title
            question: row[5] || '',     // F列: Question
            a_title: row[6] || '',      // G列: A_Title
            answer: row[7] || '',       // H列: Answer
            note: row[8] || ''          // I列: note
          });
        }
      }
      
      var result = {
        success: true,
        items: items
      };
      
      var jsonResult = JSON.stringify(result);
      
      // キャッシュに保存
      cache.put(cacheKey, jsonResult, CACHE_EXPIRY_SECONDS);
      
      return ContentService.createTextOutput(jsonResult).setMimeType(ContentService.MimeType.JSON);
      
    } catch (error) {
      // 詳細なエラー情報はログに記録（内部のみ）
      Logger.log('[Error] getCategoryData exception: ' + error.toString());
      Logger.log('[Error] Stack trace: ' + error.stack);
      // ユーザーには一般的なエラーメッセージを返す
      return createErrorResponse('Failed to retrieve category data');
    }
  }

  /**
   * Answer列を更新
   * @param {string} id - 更新対象のID（A列）
   * @param {string} answer - 更新するAnswer列の内容
   */
  function updateAnswerMemo(id, answer) {
    try {
      if (!id || id.trim() === '') {
        return createErrorResponse('ID parameter is required');
      }
      
      // スプレッドシートからデータを取得
      var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      var dataSheet = sheet.getSheetByName('Data');
      
      if (!dataSheet) {
        return createErrorResponse('Data sheet not found');
      }
      
      // A列（ID）から行を検索
      var lastRow = dataSheet.getLastRow();
      if (lastRow < 2) {
        return createErrorResponse('No data found');
      }
      
      var idRange = dataSheet.getRange(2, 1, lastRow - 1, 1); // A2からA列の最後まで
      var idValues = idRange.getValues();
      
      var targetRow = -1;
      for (var i = 0; i < idValues.length; i++) {
        var rowId = idValues[i][0];
        // 数値の場合は文字列に変換して比較
        if (typeof rowId === 'number') {
          rowId = String(rowId);
        }
        if (String(rowId) === String(id)) {
          targetRow = i + 2; // 行番号は1ベース、かつヘッダー行があるため+2
          break;
        }
      }
      
      if (targetRow === -1) {
        return createErrorResponse('ID not found');
      }
      
      // H列（Answer列）を更新
      var answerCell = dataSheet.getRange(targetRow, 8); // H列は8列目
      answerCell.setValue(answer || '');
      
      // キャッシュをクリア（該当カテゴリのキャッシュを削除）
      var cache = CacheService.getScriptCache();
      // カテゴリデータのキャッシュをクリアするため、該当行のカテゴリ番号を取得
      var categoryNo = dataSheet.getRange(targetRow, 2).getValue(); // B列は2列目
      if (categoryNo) {
        if (typeof categoryNo === 'number') {
          categoryNo = String(categoryNo);
        }
        cache.remove('categoryData_' + categoryNo);
      }
      // カテゴリ一覧のキャッシュもクリア（設問数が変わる可能性があるため）
      cache.remove('categories');
      
      var result = {
        success: true,
        message: 'Answer memo updated successfully'
      };
      
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      
    } catch (error) {
      // 詳細なエラー情報はログに記録（内部のみ）
      Logger.log('[Error] updateAnswerMemo exception: ' + error.toString());
      Logger.log('[Error] Stack trace: ' + error.stack);
      // ユーザーには一般的なエラーメッセージを返す
      return createErrorResponse('Failed to update answer memo');
    }
  }
