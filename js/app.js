// グローバル変数
var categories = [];
var currentCategoryData = [];
var currentCategoryNo = null;
var currentQuestionIndex = 0;
var learningStartTime = null;
var learningTimeInterval = null;
var stopwatchStartTime = null;
var stopwatchInterval = null;
var stopwatchElapsed = 0;
var isStopwatchRunning = false;
var isAnswerShown = false;

// 音声キャッシュ（メモリキャッシュ）
var audioCache = {};

// キャッシュの設定
var CACHE_PREFIX = 'tts_audio_'; // localStorageのキープレフィックス
var MAX_CACHE_SIZE = 10 * 1024 * 1024; // 最大キャッシュサイズ（10MB）

// スプレッドシートID
var SPREADSHEET_ID = '1pnlKMrp07Yz4MMCFByw8F04ttT3Cf6xDSX7zf5R64ZA';

// Google Sheets API v4 のAPIキー
// 注意: GitHub Pagesで公開する場合は、APIキーの制限（HTTPリファラー制限）を設定してください
// Google Cloud Console > APIとサービス > 認証情報 > APIキーを制限 > HTTPリファラー（ウェブサイト）に
// GitHub PagesのURL（例: https://yourusername.github.io/*）を追加してください
var API_KEY = 'AIzaSyCnXuzLY7ybqJU_gpl-y7gZPMO-o_7_TkY'; // ここにGoogle Cloud Consoleで取得したAPIキーを設定してください

// Google Apps Script WebアプリのURL（Cloud Text-to-Speech API用）
// 注意: Google Apps ScriptをWebアプリとして公開した際のURLを設定してください
var TTS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxnArjXdMjXLpXI36YO9JOEj6qEf2e2DDbUVfqiWdJBDB7-QwlNL3zDJS67FFmCxybMWg/exec'; // ここにGoogle Apps ScriptのWebアプリURLを設定してください

// 初期化
window.onload = function() {
  // まずカテゴリリストを読み込む（優先度：高）
  loadCategories();
  setupEventListeners();
  
  // 画像は後から読み込む（優先度：低）
  // 背景画像とボタン画像を並列で読み込む
  setBackgroundImage();
  setButtonImages();
};

// ボタン画像を設定する関数（最適化版）
function setButtonImages() {
  // ローカル画像を使用
  var images = {
    'play-button': 'img/play-button.png',
    'arrow': 'img/arrow.png',
    'home': 'img/home.png'
  };
  
  // play-button
  var playButtonImg = document.querySelector('#playButton img');
  if (playButtonImg && images['play-button']) {
    playButtonImg.src = images['play-button'];
  }
  
  // arrow (prev and next)
  var prevButtonImg = document.querySelector('#prevButton img');
  if (prevButtonImg && images['arrow']) {
    prevButtonImg.src = images['arrow'];
  }
  var nextButtonImg = document.querySelector('#nextButton img');
  if (nextButtonImg && images['arrow']) {
    nextButtonImg.src = images['arrow'];
  }
  
  // home
  var homeButtonImg = document.querySelector('#homeButton img');
  if (homeButtonImg && images['home']) {
    homeButtonImg.src = images['home'];
  }
}

// 背景画像を設定する関数（最適化版）
function setBackgroundImage() {
  var backgroundImage = document.getElementById('backgroundImage');
  if (backgroundImage) {
    // ローカル画像を使用
    backgroundImage.style.backgroundImage = 'url("img/bg.jpg")';
  }
}

// カテゴリ一覧を読み込む（最優先）
function loadCategories() {
  // ローディング表示
  var select = document.getElementById('categorySelect');
  if (select) {
    select.innerHTML = '<option value="">読み込み中...</option>';
    select.disabled = true;
  }
  
  // Google Sheets API v4でデータを取得
  // B列（Category_No）とC列（Category）を取得
  var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + SPREADSHEET_ID + '/values/Data!B2:C?key=' + API_KEY;
  
  fetch(url)
    .then(function(response) {
      if (!response.ok) {
        throw new Error('ネットワークエラー: ' + response.status);
      }
      return response.json();
    })
    .then(function(data) {
      try {
        if (!data.values || data.values.length === 0) {
          throw new Error('データがありません');
        }
        
        // カテゴリを重複排除
        var categoryMap = {};
        var categoriesList = [];
        
        for (var i = 0; i < data.values.length; i++) {
          var row = data.values[i];
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
            categoryMap[categoryNo] = true;
            categoriesList.push({
              no: categoryNo,
              name: category
            });
          }
        }
        
        if (categoriesList.length === 0) {
          throw new Error('カテゴリが見つかりません');
        }
        
        categories = categoriesList;
        if (select) {
          select.innerHTML = '<option value="">Categoryを選択してください</option>';
          categories.forEach(function(cat) {
            var option = document.createElement('option');
            option.value = cat.no;
            option.textContent = cat.no + '-' + cat.name;
            select.appendChild(option);
          });
          select.disabled = false;
        }
      } catch (e) {
        showError('データ読み込みエラー: ' + e.toString());
        if (select) {
          select.disabled = false;
        }
      }
    })
    .catch(function(error) {
      showError('アクセスエラー: ' + error.toString());
      if (select) {
        select.disabled = false;
      }
    });
}

// イベントリスナーの設定
function setupEventListeners() {
  document.getElementById('categorySelect').addEventListener('change', function() {
    var categoryNo = this.value;
    if (categoryNo) {
      // 学習時間のカウント開始（カテゴリ選択時）
      if (learningStartTime === null) {
        learningStartTime = Date.now();
        startLearningTimeCounter();
      }
      loadCategoryData(categoryNo);
    } else {
      resetListDisplay();
    }
  });
  
  document.getElementById('startButton').addEventListener('click', function() {
    startLearning();
  });
  
  document.getElementById('answerButton').addEventListener('click', function() {
    showAnswer();
  });
  
  document.getElementById('playButton').addEventListener('click', function() {
    playAnswer();
  });
  
  document.getElementById('prevButton').addEventListener('click', function() {
    goToPreviousQuestion();
  });
  
  document.getElementById('nextButton').addEventListener('click', function() {
    goToNextQuestion();
  });
  
  document.getElementById('homeButton').addEventListener('click', function() {
    goToHome();
  });
  
  // モーダル閉じるボタン
  document.getElementById('modalCloseButton').addEventListener('click', function() {
    closeModal();
  });
  
  // モーダルオーバーレイクリックで閉じる
  document.getElementById('modalOverlay').addEventListener('click', function(e) {
    if (e.target === this) {
      closeModal();
    }
  });
  
  // ESCキーでモーダルを閉じる
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      closeModal();
    }
  });
}

// カテゴリデータを読み込む
function loadCategoryData(categoryNo) {
  // Google Sheets API v4でデータを取得
  // A列（ID）からI列（note）まで取得
  var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + SPREADSHEET_ID + '/values/Data!A2:I?key=' + API_KEY;
  
  fetch(url)
    .then(function(response) {
      if (!response.ok) {
        throw new Error('ネットワークエラー: ' + response.status);
      }
      return response.json();
    })
    .then(function(data) {
      try {
        if (!data.values || data.values.length === 0) {
          throw new Error('データがありません');
        }
        
        // カテゴリ番号を文字列に統一
        if (typeof categoryNo === 'number') {
          categoryNo = String(categoryNo);
        }
        
        var items = [];
        
        for (var i = 0; i < data.values.length; i++) {
          var row = data.values[i];
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
        
        currentCategoryData = items;
        currentCategoryNo = categoryNo;
        displayList();
      } catch (e) {
        showError('データ読み込みエラー: ' + e.toString());
      }
    })
    .catch(function(error) {
      showError('アクセスエラー: ' + error.toString());
    });
}

// リストを表示
function displayList() {
  var tableBody = document.getElementById('listTableBody');
  if (!tableBody) return;
  
  tableBody.innerHTML = '';
  
  // 最初のアイテムからQ_Titleを取得してヘッダーに設定
  if (currentCategoryData.length > 0 && currentCategoryData[0].q_title) {
    var headerCell = document.getElementById('listTableHeader');
    if (headerCell) {
      headerCell.textContent = currentCategoryData[0].q_title;
    }
  }
  
  currentCategoryData.forEach(function(item) {
    var row = document.createElement('tr');
    var noCell = document.createElement('td');
    noCell.textContent = item.no || '';
    var questionCell = document.createElement('td');
    // Question列の値を表示（大文字小文字を考慮）
    var questionText = item.question || item.Question || '';
    questionCell.textContent = questionText;
    row.appendChild(noCell);
    row.appendChild(questionCell);
    
    // 行クリックでモーダルを表示
    row.addEventListener('click', function() {
      showModal(item);
    });
    
    tableBody.appendChild(row);
  });
  
  var listMessage = document.getElementById('listMessage');
  var listContainer = document.getElementById('listContainer');
  var startButton = document.getElementById('startButton');
  
  if (listMessage) listMessage.style.display = 'none';
  if (listContainer) listContainer.style.display = 'block';
  if (startButton) startButton.style.display = 'block';
}

// リスト表示をリセット
function resetListDisplay() {
  var listMessage = document.getElementById('listMessage');
  var listContainer = document.getElementById('listContainer');
  var startButton = document.getElementById('startButton');
  
  if (listMessage) listMessage.style.display = 'block';
  if (listContainer) listContainer.style.display = 'none';
  if (startButton) startButton.style.display = 'none';
}

// エラーを表示
function showError(message) {
  var errorDiv = document.createElement('div');
  errorDiv.className = 'error-message';
  errorDiv.textContent = message;
  var container = document.querySelector('.container');
  if (container) {
    container.insertBefore(errorDiv, container.firstChild);
    setTimeout(function() {
      errorDiv.remove();
    }, 3000);
  }
}

// 学習開始
function startLearning() {
  if (currentCategoryData.length === 0) {
    return;
  }
  
  // 画面遷移
  var screen1 = document.getElementById('screen1');
  var screen2 = document.getElementById('screen2');
  if (screen1) screen1.classList.remove('active');
  if (screen2) screen2.classList.add('active');
  
  // コンテナのパディングを減らす
  var container = document.querySelector('.container');
  if (container) container.classList.add('learning-mode');
  
  // カテゴリ情報を表示
  var selectedCategory = categories.find(function(cat) {
    return cat.no == currentCategoryNo;
  });
  if (selectedCategory) {
    var currentCategory = document.getElementById('currentCategory');
    if (currentCategory) {
      currentCategory.textContent = selectedCategory.no + '. ' + selectedCategory.name;
    }
  }
  
  // 最初の問題を表示
  currentQuestionIndex = 0;
  displayQuestion();
  
  // 最初の問題と次の問題をプリロード
  preloadAudioForCurrentAndNext();
}

// 学習時間カウンターを開始
function startLearningTimeCounter() {
  learningTimeInterval = setInterval(function() {
    updateLearningTime();
  }, 1000);
  updateLearningTime();
}

// 学習時間を更新
function updateLearningTime() {
  if (learningStartTime === null) return;
  
  var elapsed = Date.now() - learningStartTime;
  var totalMinutes = Math.floor(elapsed / 60000);
  var hours = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  var seconds = Math.floor((elapsed / 1000) % 60);
  
  var timeText = '<学習時間>' + hours + '時間' + String(minutes).padStart(2, '0') + '分' + String(seconds).padStart(2, '0') + '秒';
  
  // 学習画面の学習時間を更新
  var learningTimeElement = document.getElementById('learningTime');
  if (learningTimeElement) {
    learningTimeElement.textContent = timeText;
  }
  
  // TOPページの学習時間を更新
  var learningTimeTopElement = document.getElementById('learningTimeTop');
  if (learningTimeTopElement) {
    learningTimeTopElement.textContent = timeText;
  }
}

// 問題を表示
function displayQuestion() {
  if (currentQuestionIndex < 0 || currentQuestionIndex >= currentCategoryData.length) {
    return;
  }
  
  var item = currentCategoryData[currentQuestionIndex];
  
  // セクションラベルをQ_TitleとA_Titleから設定
  if (item.q_title) {
    var questionLabel = document.getElementById('questionSectionLabel');
    if (questionLabel) {
      questionLabel.textContent = item.q_title;
    }
  }
  if (item.a_title) {
    var answerLabel = document.getElementById('answerSectionLabel');
    if (answerLabel) {
      answerLabel.textContent = item.a_title;
    }
  }
  
  // 出題数表示
  var questionInfo = document.getElementById('questionInfo');
  if (questionInfo) {
    questionInfo.textContent = '[' + (currentQuestionIndex + 1) + '/' + currentCategoryData.length + ']';
  }
  
  // 質問文を表示
  var questionText = document.getElementById('questionText');
  if (questionText) {
    questionText.textContent = item.question || '';
  }
  
  // Answerボタンを表示
  var answerButtonContainer = document.getElementById('answerButtonContainer');
  var answerText = document.getElementById('answerText');
  if (answerButtonContainer) answerButtonContainer.style.display = 'block';
  if (answerText) answerText.classList.add('blinking');
  
  // 回答テキスト、停止時間表示コンテナを非表示
  var answerTextDisplay = document.getElementById('answerTextDisplay');
  var answerTimeContainer = document.getElementById('answerTimeContainer');
  var noteSection = document.getElementById('noteSection');
  if (answerTextDisplay) answerTextDisplay.style.display = 'none';
  if (answerTimeContainer) answerTimeContainer.style.display = 'none';
  if (noteSection) noteSection.style.display = 'none';
  isAnswerShown = false;
  
  // ストップウォッチをリセットして開始
  resetStopwatch();
  startStopwatch();
  
  // ナビゲーションボタンを無効化（Answerボタンが押されるまで）
  var prevButton = document.getElementById('prevButton');
  var nextButton = document.getElementById('nextButton');
  if (prevButton) prevButton.disabled = true;
  if (nextButton) nextButton.disabled = true;
  
  // 次の問題をプリロード（バックグラウンドで非同期実行）
  preloadNextQuestions();
}

// ストップウォッチを開始
function startStopwatch() {
  if (isStopwatchRunning) return;
  
  stopwatchStartTime = Date.now() - stopwatchElapsed;
  isStopwatchRunning = true;
  stopwatchInterval = setInterval(function() {
    updateStopwatch();
  }, 10);
  updateStopwatch();
}

// ストップウォッチを停止
function stopStopwatch() {
  if (!isStopwatchRunning) return;
  
  clearInterval(stopwatchInterval);
  stopwatchElapsed = Date.now() - stopwatchStartTime;
  isStopwatchRunning = false;
}

// ストップウォッチをリセット
function resetStopwatch() {
  stopStopwatch();
  stopwatchElapsed = 0;
  var answerStopwatch = document.getElementById('answerStopwatch');
  if (answerStopwatch) {
    answerStopwatch.textContent = '(00:00:00)';
  }
}

// ストップウォッチを更新
function updateStopwatch() {
  if (!isStopwatchRunning) return;
  
  var elapsed = Date.now() - stopwatchStartTime;
  var totalSeconds = Math.floor(elapsed / 1000);
  var minutes = Math.floor(totalSeconds / 60);
  var seconds = totalSeconds % 60;
  var milliseconds = Math.floor((elapsed % 1000) / 10);
  
  var answerStopwatch = document.getElementById('answerStopwatch');
  if (answerStopwatch) {
    answerStopwatch.textContent = 
      '(' + String(minutes).padStart(2, '0') + ':' +
      String(seconds).padStart(2, '0') + ':' +
      String(milliseconds).padStart(2, '0') + ')';
  }
}

// 答えを表示
function showAnswer() {
  if (isAnswerShown) return;
  
  stopStopwatch();
  
  var item = currentCategoryData[currentQuestionIndex];
  
  // Answerボタンを非表示
  var answerButtonContainer = document.getElementById('answerButtonContainer');
  var answerText = document.getElementById('answerText');
  if (answerButtonContainer) answerButtonContainer.style.display = 'none';
  if (answerText) answerText.classList.remove('blinking');
  
  // 停止時間を計算して表示
  var elapsed = stopwatchElapsed;
  var totalSeconds = Math.floor(elapsed / 1000);
  var minutes = Math.floor(totalSeconds / 60);
  var seconds = totalSeconds % 60;
  var milliseconds = Math.floor((elapsed % 1000) / 10);
  var timeText = String(minutes).padStart(2, '0') + ':' +
                 String(seconds).padStart(2, '0') + ':' +
                 String(milliseconds).padStart(2, '0');
  var answerTimeDisplay = document.getElementById('answerTimeDisplay');
  if (answerTimeDisplay) {
    answerTimeDisplay.textContent = timeText;
  }
  
  // 停止時間表示コンテナを表示（停止時間とplayボタンを含む）
  var answerTimeContainer = document.getElementById('answerTimeContainer');
  if (answerTimeContainer) {
    answerTimeContainer.style.display = 'flex';
  }
  
  // 回答文を表示
  var answerTextDisplay = document.getElementById('answerTextDisplay');
  if (answerTextDisplay) {
    answerTextDisplay.textContent = item.answer || '';
    answerTextDisplay.classList.remove('answer-hidden');
    answerTextDisplay.style.display = 'block';
  }
  
  // noteを表示
  if (item.note) {
    var noteText = document.getElementById('noteText');
    var noteSection = document.getElementById('noteSection');
    if (noteText) noteText.textContent = item.note;
    if (noteSection) noteSection.style.display = 'block';
  }
  
  isAnswerShown = true;
  
  // ナビゲーションボタンを有効化
  updateNavigationButtons();
}

// 回答を読み上げ
function playAnswer() {
  var item = currentCategoryData[currentQuestionIndex];
  if (!item || !item.answer) return;
  
  // WebアプリURLが設定されていない場合はエラー
  if (!TTS_WEB_APP_URL || TTS_WEB_APP_URL === 'YOUR_WEB_APP_URL_HERE') {
    showError('音声読み上げの設定が完了していません。WebアプリURLを設定してください。');
    return;
  }
  
  var text = item.answer;
  
  // キャッシュから音声データを取得
  var cachedAudio = getCachedAudio(text);
  if (cachedAudio) {
    // キャッシュから即座に再生
    playAudioFromCache(cachedAudio);
    return;
  }
  
  // キャッシュにない場合はAPI呼び出し
  fetchAudioFromAPI(text);
}

/**
 * キャッシュから音声データを取得
 * メモリキャッシュ → localStorage の順で確認
 */
function getCachedAudio(text) {
  // メモリキャッシュを確認
  if (audioCache[text]) {
    return audioCache[text];
  }
  
  // localStorageを確認
  try {
    var cacheKey = CACHE_PREFIX + hashText(text);
    var cachedData = localStorage.getItem(cacheKey);
    if (cachedData) {
      var audioData = JSON.parse(cachedData);
      // メモリキャッシュにも保存
      audioCache[text] = audioData;
      return audioData;
    }
  } catch (e) {
    // localStorageが使用できない場合やエラーが発生した場合は無視
    console.warn('Cache read error:', e);
  }
  
  return null;
}

/**
 * 音声データをキャッシュに保存
 */
function saveAudioToCache(text, audioContent) {
  var audioData = {
    audioContent: audioContent,
    timestamp: Date.now()
  };
  
  // メモリキャッシュに保存
  audioCache[text] = audioData;
  
  // localStorageに保存（サイズ制限を考慮）
  try {
    var cacheKey = CACHE_PREFIX + hashText(text);
    var dataToStore = JSON.stringify(audioData);
    
    // キャッシュサイズをチェック
    if (getCacheSize() + dataToStore.length > MAX_CACHE_SIZE) {
      // キャッシュが大きすぎる場合は古いエントリを削除
      clearOldCacheEntries();
    }
    
    localStorage.setItem(cacheKey, dataToStore);
  } catch (e) {
    // localStorageが満杯の場合やエラーが発生した場合は無視
    console.warn('Cache save error:', e);
    // 古いキャッシュを削除して再試行
    try {
      clearOldCacheEntries();
      localStorage.setItem(cacheKey, JSON.stringify(audioData));
    } catch (e2) {
      // それでも失敗した場合はメモリキャッシュのみ使用
      console.warn('Cache save retry failed:', e2);
    }
  }
}

/**
 * テキストをハッシュ化（localStorageのキー用）
 */
function hashText(text) {
  var hash = 0;
  for (var i = 0; i < text.length; i++) {
    var char = text.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return Math.abs(hash).toString(36);
}

/**
 * 現在のキャッシュサイズを取得
 */
function getCacheSize() {
  var totalSize = 0;
  try {
    for (var i = 0; i < localStorage.length; i++) {
      var key = localStorage.key(i);
      if (key && key.indexOf(CACHE_PREFIX) === 0) {
        var value = localStorage.getItem(key);
        if (value) {
          totalSize += value.length;
        }
      }
    }
  } catch (e) {
    // エラーが発生した場合は0を返す
  }
  return totalSize;
}

/**
 * 古いキャッシュエントリを削除（FIFO方式）
 */
function clearOldCacheEntries() {
  try {
    var entries = [];
    for (var i = 0; i < localStorage.length; i++) {
      var key = localStorage.key(i);
      if (key && key.indexOf(CACHE_PREFIX) === 0) {
        var value = localStorage.getItem(key);
        if (value) {
          try {
            var data = JSON.parse(value);
            entries.push({
              key: key,
              timestamp: data.timestamp || 0
            });
          } catch (e) {
            // パースエラーは無視
          }
        }
      }
    }
    
    // タイムスタンプでソート（古い順）
    entries.sort(function(a, b) {
      return a.timestamp - b.timestamp;
    });
    
    // 古いエントリの50%を削除
    var deleteCount = Math.floor(entries.length / 2);
    for (var j = 0; j < deleteCount; j++) {
      localStorage.removeItem(entries[j].key);
      // メモリキャッシュからも削除（該当するものがあれば）
      for (var text in audioCache) {
        if (hashText(text) === entries[j].key.replace(CACHE_PREFIX, '')) {
          delete audioCache[text];
        }
      }
    }
  } catch (e) {
    console.warn('Cache clear error:', e);
  }
}

/**
 * キャッシュから音声を再生
 */
function playAudioFromCache(audioData) {
  if (!audioData || !audioData.audioContent) {
    return;
  }
  
  try {
    var audio = new Audio('data:audio/mp3;base64,' + audioData.audioContent);
    audio.play().catch(function(error) {
      showError('音声の再生に失敗しました: ' + error.toString());
    });
  } catch (error) {
    showError('音声の再生に失敗しました: ' + error.toString());
  }
}

/**
 * ローディング表示を開始
 */
function showPlayButtonLoading() {
  var playButton = document.getElementById('playButton');
  if (playButton) {
    playButton.disabled = true;
    // スピナーを表示
    var spinner = document.createElement('div');
    spinner.className = 'play-button-spinner';
    spinner.id = 'playButtonSpinner';
    playButton.innerHTML = '';
    playButton.appendChild(spinner);
  }
}

/**
 * ローディング表示を終了
 */
function hidePlayButtonLoading() {
  var playButton = document.getElementById('playButton');
  if (playButton) {
    playButton.disabled = false;
    // 元の画像を復元
    var playButtonImg = document.createElement('img');
    playButtonImg.src = 'img/play-button.png';
    playButtonImg.alt = '再生';
    playButton.innerHTML = '';
    playButton.appendChild(playButtonImg);
  }
}

/**
 * APIから音声データを取得
 */
function fetchAudioFromAPI(text) {
  // ローディング表示を開始
  showPlayButtonLoading();
  
  var playButton = document.getElementById('playButton');
  
  // リクエストパラメータを準備
  var params = new URLSearchParams();
  params.append('text', text);
  params.append('referer', window.location.origin);
  
  // Google Apps Scriptにリクエストを送信
  fetch(TTS_WEB_APP_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params
  })
  .then(function(response) {
    if (!response.ok) {
      throw new Error('ネットワークエラー: ' + response.status);
    }
    return response.json();
  })
  .then(function(data) {
    // ローディング表示を終了
    hidePlayButtonLoading();
    
    if (data.success && data.audioContent) {
      // キャッシュに保存
      saveAudioToCache(text, data.audioContent);
      
      // 音声データ（base64）を再生
      var audio = new Audio('data:audio/mp3;base64,' + data.audioContent);
      audio.play().catch(function(error) {
        showError('音声の再生に失敗しました: ' + error.toString());
      });
    } else {
      showError('音声の生成に失敗しました: ' + (data.error || 'Unknown error'));
    }
  })
  .catch(function(error) {
    // ローディング表示を終了
    hidePlayButtonLoading();
    showError('音声読み上げエラー: ' + error.toString());
  });
}

/**
 * 現在の問題と次の問題の音声をプリロード
 */
function preloadAudioForCurrentAndNext() {
  if (!TTS_WEB_APP_URL || TTS_WEB_APP_URL === 'YOUR_WEB_APP_URL_HERE') {
    return; // WebアプリURLが設定されていない場合はスキップ
  }
  
  // 現在の問題（最初の問題）をプリロード
  if (currentQuestionIndex >= 0 && currentQuestionIndex < currentCategoryData.length) {
    var currentItem = currentCategoryData[currentQuestionIndex];
    if (currentItem && currentItem.answer) {
      preloadAudio(currentItem.answer);
    }
  }
  
  // 次の問題をプリロード
  preloadNextQuestions();
}

/**
 * 次の問題（最大2問）の音声をプリロード
 */
function preloadNextQuestions() {
  if (!TTS_WEB_APP_URL || TTS_WEB_APP_URL === 'YOUR_WEB_APP_URL_HERE') {
    return; // WebアプリURLが設定されていない場合はスキップ
  }
  
  var preloadCount = 2; // 次の2問をプリロード
  
  for (var i = 1; i <= preloadCount; i++) {
    var nextIndex = currentQuestionIndex + i;
    if (nextIndex >= 0 && nextIndex < currentCategoryData.length) {
      var nextItem = currentCategoryData[nextIndex];
      if (nextItem && nextItem.answer) {
        preloadAudio(nextItem.answer);
      }
    }
  }
}

/**
 * 指定されたテキストの音声をプリロード（バックグラウンドで非同期実行）
 */
function preloadAudio(text) {
  if (!text || !text.trim()) {
    return;
  }
  
  // キャッシュに既に存在する場合はスキップ
  var cachedAudio = getCachedAudio(text);
  if (cachedAudio) {
    return; // 既にキャッシュされている
  }
  
  // バックグラウンドで非同期にプリロード（エラーは無視）
  setTimeout(function() {
    var params = new URLSearchParams();
    params.append('text', text);
    params.append('referer', window.location.origin);
    
    fetch(TTS_WEB_APP_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params
    })
    .then(function(response) {
      if (!response.ok) {
        return; // エラーは無視
      }
      return response.json();
    })
    .then(function(data) {
      if (data && data.success && data.audioContent) {
        // キャッシュに保存（再生はしない）
        saveAudioToCache(text, data.audioContent);
      }
    })
    .catch(function(error) {
      // プリロードのエラーは無視（ユーザーに影響を与えない）
      // console.log('Preload error (ignored):', error);
    });
  }, 100); // 少し遅延させて、メイン処理を優先
}

// 前の問題に戻る
function goToPreviousQuestion() {
  if (currentQuestionIndex > 0) {
    currentQuestionIndex--;
    displayQuestion();
  }
}

// 次の問題に進む
function goToNextQuestion() {
  if (currentQuestionIndex < currentCategoryData.length - 1) {
    currentQuestionIndex++;
    displayQuestion();
  }
}

// ナビゲーションボタンの状態を更新
function updateNavigationButtons() {
  var prevButton = document.getElementById('prevButton');
  var nextButton = document.getElementById('nextButton');
  if (prevButton) prevButton.disabled = (currentQuestionIndex === 0);
  if (nextButton) nextButton.disabled = (currentQuestionIndex === currentCategoryData.length - 1);
}

// ホームに戻る
function goToHome() {
  // ストップウォッチを停止
  stopStopwatch();
  
  // 画面遷移
  var screen2 = document.getElementById('screen2');
  var screen1 = document.getElementById('screen1');
  if (screen2) screen2.classList.remove('active');
  if (screen1) screen1.classList.add('active');
  
  // コンテナのパディングを元に戻す
  var container = document.querySelector('.container');
  if (container) container.classList.remove('learning-mode');
  
  // 学習時間はリセットしない（継続）
}

// モーダルを表示
function showModal(item) {
  if (!item) return;
  
  // A_Titleをラベルに設定
  if (item.a_title) {
    var modalAnswerLabel = document.getElementById('modalAnswerLabel');
    if (modalAnswerLabel) {
      modalAnswerLabel.textContent = item.a_title;
    }
  }
  
  // 回答文を表示
  var answerText = document.getElementById('modalAnswerText');
  if (answerText && item.answer) {
    answerText.textContent = item.answer;
  }
  
  // noteを表示（ある場合のみ）
  var noteSection = document.getElementById('modalNoteSection');
  var noteText = document.getElementById('modalNoteText');
  if (item.note && item.note.trim()) {
    if (noteText) {
      noteText.textContent = item.note;
    }
    if (noteSection) {
      noteSection.style.display = 'block';
    }
  } else {
    if (noteSection) {
      noteSection.style.display = 'none';
    }
  }
  
  // モーダルを表示
  var modalOverlay = document.getElementById('modalOverlay');
  if (modalOverlay) {
    modalOverlay.classList.add('active');
  }
}

// モーダルを閉じる
function closeModal() {
  var modalOverlay = document.getElementById('modalOverlay');
  if (modalOverlay) {
    modalOverlay.classList.remove('active');
  }
}

