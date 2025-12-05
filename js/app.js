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
var userEmail = null; // ユーザーのメールアドレス
var modalCurrentIndex = 0; // モーダル内の現在のインデックス
var retryQuestionIndices = []; // 再チャレンジする問題のインデックスを保存
var isInRetryMode = false; // 再チャレンジモードかどうか
var retryQuestionIndex = 0; // 現在の再チャレンジ問題のインデックス
var completedQuestionIndices = []; // 完了した問題のインデックスを保存（灰色表示）
var isLearningCompleted = false; // 学習が完了したかどうか
var selectedQuestionIndices = []; // 選択された問題のインデックスを保存
var originalCategoryData = []; // 元の全問題データ（出題数表示用）
var isToggleActive = true; // トグルボタンの状態（ON/OFF）
var currentAudio = null; // 現在再生中のAudioオブジェクト

// 音声キャッシュ（メモリキャッシュ）
var audioCache = {};

// キャッシュの設定
var CACHE_PREFIX = 'tts_audio_'; // localStorageのキープレフィックス
var MAX_CACHE_SIZE = 10 * 1024 * 1024; // 最大キャッシュサイズ（10MB）

// Google Apps Script WebアプリのURL（統合版：TTSとDATAの両方を処理）
// 注意: Gas_Main.gsをWebアプリとして公開した際のURLを設定してください
var WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxTBkXrUOsYjzb1xERU-GXe5g8w9f0lxqOyxn6P8-VC9zNDMtjmTXOKRH_lBnRra3Kzcw/exec'; // ここにGoogle Apps ScriptのWebアプリURLを設定してください

// 学習完了メッセージの定数配列
var COMPLETION_MESSAGES = [
  'Good job!',
  'がんばってるじゃん！',
  'Excellent!',
  'その調子！'
];

// 学習完了メッセージ用のアイコン画像ファイル名
var COMPLETION_MESSAGE_IMAGES = [
  'msg-ino-01.png',
  'msg-manmos-01.png',
  'msg-putera-01.png',
  'msg-smiley-01.png',
  'msg-smiley-02.png',
  'msg-thumb-01.png',
  'msg-uma-01.png'
];

// 初期化
window.onload = function() {
  // メールアドレスを確認
  checkUserEmail();
  
  setupEventListeners();
  
  // 画像は後から読み込む（優先度：低）
  // 背景画像とボタン画像を並列で読み込む
  setBackgroundImage();
  setButtonImages();
  
  // トグルボタンの初期状態を設定
  var toggleButton = document.getElementById('toggleButton');
  if (toggleButton) {
    if (isToggleActive) {
      toggleButton.classList.add('active');
    } else {
      toggleButton.classList.remove('active');
    }
  }
  
  // トグルボタンの位置を設定（タイトルが表示された後に実行）
  requestAnimationFrame(function() {
    requestAnimationFrame(updateToggleButtonPosition);
  });
  
  // ウィンドウリサイズ時にも位置を更新
  window.addEventListener('resize', updateToggleButtonPosition);
};

// ページローディングを非表示にする
function hidePageLoading() {
  var loadingOverlay = document.getElementById('pageLoadingOverlay');
  if (loadingOverlay) {
    // フェードアウトアニメーション
    loadingOverlay.classList.add('hidden');
    // アニメーション完了後にDOMから削除
    setTimeout(function() {
      if (loadingOverlay.parentNode) {
        loadingOverlay.parentNode.removeChild(loadingOverlay);
      }
    }, 300); // transition時間（0.3s）に合わせる
  }
}

// メールアドレスを確認し、必要に応じて入力画面を表示
function checkUserEmail() {
  // localStorageからメールアドレスを取得
  userEmail = localStorage.getItem('userEmail');
  
  if (!userEmail) {
    // メールアドレスが保存されていない場合は入力画面を表示
    showEmailInputDialog();
  } else {
    // メールアドレスが保存されている場合はカテゴリリストを読み込む
    loadCategories();
  }
}

// メールアドレス入力ダイアログを表示
function showEmailInputDialog() {
  var email = prompt('メールアドレスを入力してください:');
  
  // nullの場合はキャンセルが押された
  if (email === null) {
    return; // 何もせずに終了
  }
  
  if (email && email.trim() !== '') {
    userEmail = email.trim();
    // localStorageに保存
    localStorage.setItem('userEmail', userEmail);
    
    // ログイン成功時はエラーメッセージを自動削除
    clearErrorMessages();
    
    // カテゴリリストを読み込む
    loadCategories();
  } else {
    // メールアドレスが入力されなかった場合は再度表示
    alert('メールアドレスは必須です。');
    showEmailInputDialog();
  }
}

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
  
  // arrow (next)
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
  // userEmailが設定されていない場合は、再度確認
  if (!userEmail) {
    userEmail = localStorage.getItem('userEmail');
  }
  
  if (!userEmail) {
    showError('メールアドレスが設定されていません。');
    checkUserEmail();
    return;
  }
  
  // ローディング表示
  var select = document.getElementById('categorySelect');
  var loadingSpinner = document.getElementById('categoryLoadingSpinner');
  if (select) {
    select.innerHTML = '<option value="">読み込み中...</option>';
    select.disabled = true;
  }
  if (loadingSpinner) {
    loadingSpinner.style.display = 'block';
  }
  
  // Google Apps Script経由でデータを取得
  var params = new URLSearchParams();
  params.append('action', 'getCategories');
  params.append('email', userEmail);
  params.append('referer', window.location.origin);
  
  // GETリクエストで送信
  var requestUrl = WEB_APP_URL + '?' + params.toString();
  
  fetch(requestUrl)
    .then(function(response) {
      if (!response.ok) {
        throw new Error('ネットワークエラー: ' + response.status);
      }
      return response.json();
    })
    .then(function(data) {
      try {
        if (!data.success) {
          throw new Error(data.error || 'データの取得に失敗しました');
        }
        
        if (!data.categories || data.categories.length === 0) {
          throw new Error('カテゴリが見つかりません');
        }
        
        categories = data.categories;
        if (select) {
          select.innerHTML = '<option value="">Categoryを選択してください</option>';
          categories.forEach(function(cat) {
            var option = document.createElement('option');
            option.value = cat.no;
            // 設問数を表示（countが存在する場合のみ）
            var displayText = cat.no + '-' + cat.name;
            if (cat.count !== undefined && cat.count !== null) {
              displayText += '（' + cat.count + '問）';
            }
            option.textContent = displayText;
            select.appendChild(option);
          });
          select.disabled = false;
        }
        if (loadingSpinner) {
          loadingSpinner.style.display = 'none';
        }
        // ボタンの状態を更新
        updateListNavButtons();
        // ページローディングを非表示（Googleスプレッドシートの読み込み完了）
        hidePageLoading();
      } catch (e) {
        showError('データ読み込みエラー: ' + e.toString());
        if (select) {
          select.disabled = false;
        }
        if (loadingSpinner) {
          loadingSpinner.style.display = 'none';
        }
        // ページローディングを非表示（エラー時も非表示）
        hidePageLoading();
      }
    })
    .catch(function(error) {
      showError('アクセスエラー: ' + error.toString());
      if (select) {
        select.disabled = false;
      }
      if (loadingSpinner) {
        loadingSpinner.style.display = 'none';
      }
      // ページローディングを非表示（エラー時も非表示）
      hidePageLoading();
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
    // ボタンの状態はloadCategoryData()内で更新されるため、ここでは呼び出さない
  });
  
  document.getElementById('startButton').addEventListener('click', function() {
    startLearning();
  });
  
  // 旧Answerボタン（削除済み）の代わりに、ナビゲーションバーのAnswerボタンを使用
  document.getElementById('navAnswerButton').addEventListener('click', function() {
    showAnswer();
  });
  
  document.getElementById('playButton').addEventListener('click', function() {
    playAnswer();
  });
  
  document.getElementById('nextButton').addEventListener('click', function() {
    goToNextQuestion();
  });
  
  document.getElementById('homeButton').addEventListener('click', function() {
    goToHome();
  });
  
  document.getElementById('plusButton').addEventListener('click', function() {
    handlePlusButtonClick();
  });
  
  document.getElementById('toggleButton').addEventListener('click', function() {
    isToggleActive = !isToggleActive;
    if (isToggleActive) {
      this.classList.add('active');
    } else {
      this.classList.remove('active');
    }
  });
  
  document.getElementById('loginButton').addEventListener('click', function() {
    showEmailInputDialog();
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
  
  // モーダル内の前へボタン
  document.getElementById('modalPrevButton').addEventListener('click', function() {
    if (modalCurrentIndex > 0) {
      modalCurrentIndex--;
      var item = currentCategoryData[modalCurrentIndex];
      if (item) {
        updateModalContent(item);
        updateModalNavigation();
        updateModalSelection();
      }
    }
  });
  
  // モーダル内の次へボタン
  document.getElementById('modalNextButton').addEventListener('click', function() {
    if (modalCurrentIndex < currentCategoryData.length - 1) {
      modalCurrentIndex++;
      var item = currentCategoryData[modalCurrentIndex];
      if (item) {
        updateModalContent(item);
        updateModalNavigation();
        updateModalSelection();
      }
    }
  });
  
  // モーダル内の選択ボタン
  document.getElementById('modalSelectButton').addEventListener('click', function() {
    handleModalSelection();
  });
  
  // クリアボタン
  document.getElementById('clearSelectionButton').addEventListener('click', function() {
    clearSelection();
  });
  
  // Listナビゲーションボタン（前へ）
  document.getElementById('listPrevButton').addEventListener('click', function() {
    navigateToPreviousCategory();
  });
  
  // Listナビゲーションボタン（次へ）
  document.getElementById('listNextButton').addEventListener('click', function() {
    navigateToNextCategory();
  });
}

// カテゴリデータを読み込む
function loadCategoryData(categoryNo) {
  // userEmailが設定されていない場合は、再度確認
  if (!userEmail) {
    userEmail = localStorage.getItem('userEmail');
  }
  
  if (!userEmail) {
    showError('メールアドレスが設定されていません。');
    checkUserEmail();
    return;
  }
  
  // ローディング表示
  var loadingSpinner = document.getElementById('categoryLoadingSpinner');
  if (loadingSpinner) {
    loadingSpinner.style.display = 'block';
  }
  
  // ボタンを無効化
  var prevButton = document.getElementById('listPrevButton');
  var nextButton = document.getElementById('listNextButton');
  if (prevButton) prevButton.disabled = true;
  if (nextButton) nextButton.disabled = true;
  
  // Google Apps Script経由でデータを取得
  var params = new URLSearchParams();
  params.append('action', 'getCategoryData');
  params.append('categoryNo', categoryNo);
  params.append('email', userEmail);
  params.append('referer', window.location.origin);
  
  // GETリクエストで送信
  var requestUrl = WEB_APP_URL + '?' + params.toString();
  
  fetch(requestUrl)
    .then(function(response) {
      if (!response.ok) {
        throw new Error('ネットワークエラー: ' + response.status);
      }
      return response.json();
    })
    .then(function(data) {
      try {
        if (!data.success) {
          throw new Error(data.error || 'データの取得に失敗しました');
        }
        
        if (!data.items) {
          throw new Error('データがありません');
        }
        
        currentCategoryData = data.items;
        currentCategoryNo = categoryNo;
        // 選択状態をリセット
        selectedQuestionIndices = [];
        displayList();
        // ボタンの状態を更新（表示/非表示と有効/無効を設定）
        updateListNavButtons();
        
        // ローディング非表示
        if (loadingSpinner) {
          loadingSpinner.style.display = 'none';
        }
      } catch (e) {
        showError('データ読み込みエラー: ' + e.toString());
        // ローディング非表示
        if (loadingSpinner) {
          loadingSpinner.style.display = 'none';
        }
        // ボタンの状態を更新（エラー時も無効化のまま）
        updateListNavButtons();
      }
    })
    .catch(function(error) {
      showError('アクセスエラー: ' + error.toString());
      // ローディング非表示
      if (loadingSpinner) {
        loadingSpinner.style.display = 'none';
      }
      // ボタンの状態を更新（エラー時も無効化のまま）
      updateListNavButtons();
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
  
  currentCategoryData.forEach(function(item, index) {
    var row = document.createElement('tr');
    var isSelected = selectedQuestionIndices.indexOf(index) !== -1;
    
    // 選択状態に応じてクラスを追加
    if (isSelected) {
      row.classList.add('selected-row');
    }
    
    var noCell = document.createElement('td');
    noCell.textContent = item.no || '';
    // 選択状態に応じてNo列にクラスを追加
    if (isSelected) {
      noCell.classList.add('selected-no');
    }
    var questionCell = document.createElement('td');
    // Question列の値を表示（画像対応）
    var questionContent = item.question || item.Question || '';
    if (isImageUrl(questionContent)) {
      // 画像URLの場合はサムネイル表示
      var imageUrl = convertGoogleDriveUrl(questionContent);
      var img = document.createElement('img');
      img.src = imageUrl;
      img.className = 'list-thumbnail';
      img.alt = '画像';
      img.style.maxWidth = '100px';
      img.style.maxHeight = '60px';
      img.style.height = 'auto';
      img.style.display = 'block';
      img.style.objectFit = 'contain';
      
      // エラーハンドリング
      img.addEventListener('error', function() {
        questionCell.textContent = '[画像]';
      });
      
      questionCell.appendChild(img);
    } else {
      // テキストの場合はテキスト表示
      questionCell.textContent = questionContent;
    }
    row.appendChild(noCell);
    row.appendChild(questionCell);
    
    // シングルクリックで選択/解除（トグル）
    var clickTimer = null;
    row.addEventListener('click', function(e) {
      if (clickTimer === null) {
        clickTimer = setTimeout(function() {
          clickTimer = null;
          // シングルクリック：選択/解除
          toggleQuestionSelection(index, row);
        }, 300);
      }
    });
    
    // ダブルクリックでモーダル表示
    row.addEventListener('dblclick', function(e) {
      e.preventDefault();
      if (clickTimer) {
        clearTimeout(clickTimer);
        clickTimer = null;
      }
      var itemIndex = currentCategoryData.indexOf(item);
      showModal(item, itemIndex);
    });
    
    tableBody.appendChild(row);
  });
  
  // 選択数の表示を更新
  updateSelectionCount();
  
  var listMessage = document.getElementById('listMessage');
  var listContainer = document.getElementById('listContainer');
  var startButton = document.getElementById('startButton');
  
  if (listMessage) listMessage.style.display = 'none';
  if (listContainer) listContainer.style.display = 'block';
  if (startButton) startButton.style.display = 'block';
}

// 問題の選択/解除をトグル
function toggleQuestionSelection(index, row) {
  var selectedIndex = selectedQuestionIndices.indexOf(index);
  var noCell = row.querySelector('td:first-child'); // No列を取得
  if (selectedIndex === -1) {
    // 選択
    selectedQuestionIndices.push(index);
    row.classList.add('selected-row');
    if (noCell) noCell.classList.add('selected-no');
  } else {
    // 解除
    selectedQuestionIndices.splice(selectedIndex, 1);
    row.classList.remove('selected-row');
    if (noCell) noCell.classList.remove('selected-no');
  }
  updateSelectionCount();
}

// Listナビゲーションボタンを表示
function showListNavButtons() {
  var listNavContainer = document.querySelector('.list-nav-container');
  if (listNavContainer) {
    listNavContainer.style.visibility = 'visible';
  }
}

// Listナビゲーションボタンを非表示
function hideListNavButtons() {
  var listNavContainer = document.querySelector('.list-nav-container');
  if (listNavContainer) {
    listNavContainer.style.visibility = 'hidden';
  }
}

// 前のカテゴリに移動
function navigateToPreviousCategory() {
  var select = document.getElementById('categorySelect');
  if (!select || !select.value || categories.length === 0) {
    return;
  }
  
  // ボタンを無効化
  var prevButton = document.getElementById('listPrevButton');
  var nextButton = document.getElementById('listNextButton');
  if (prevButton) prevButton.disabled = true;
  if (nextButton) nextButton.disabled = true;
  
  // 現在選択されているカテゴリのインデックスを取得
  var currentIndex = -1;
  for (var i = 0; i < categories.length; i++) {
    if (categories[i].no == select.value) {
      currentIndex = i;
      break;
    }
  }
  
  // 前のカテゴリが存在する場合
  if (currentIndex > 0) {
    var previousCategory = categories[currentIndex - 1];
    select.value = previousCategory.no;
    // changeイベントを手動で発火
    var event = new Event('change', { bubbles: true });
    select.dispatchEvent(event);
  }
}

// 次のカテゴリに移動
function navigateToNextCategory() {
  var select = document.getElementById('categorySelect');
  if (!select || !select.value || categories.length === 0) {
    return;
  }
  
  // ボタンを無効化
  var prevButton = document.getElementById('listPrevButton');
  var nextButton = document.getElementById('listNextButton');
  if (prevButton) prevButton.disabled = true;
  if (nextButton) nextButton.disabled = true;
  
  // 現在選択されているカテゴリのインデックスを取得
  var currentIndex = -1;
  for (var i = 0; i < categories.length; i++) {
    if (categories[i].no == select.value) {
      currentIndex = i;
      break;
    }
  }
  
  // 次のカテゴリが存在する場合
  if (currentIndex >= 0 && currentIndex < categories.length - 1) {
    var nextCategory = categories[currentIndex + 1];
    select.value = nextCategory.no;
    // changeイベントを手動で発火
    var event = new Event('change', { bubbles: true });
    select.dispatchEvent(event);
  }
}

// Listナビゲーションボタンの状態を更新
function updateListNavButtons() {
  var prevButton = document.getElementById('listPrevButton');
  var nextButton = document.getElementById('listNextButton');
  var select = document.getElementById('categorySelect');
  
  if (!prevButton || !nextButton || !select || categories.length === 0) {
    // カテゴリが読み込まれていない場合は非表示
    hideListNavButtons();
    return;
  }
  
  // カテゴリが選択されていない場合
  if (!select.value) {
    // 非表示にする
    hideListNavButtons();
    return;
  }
  
  // カテゴリが選択されている場合は表示
  showListNavButtons();
  
  // 現在選択されているカテゴリのインデックスを取得
  var currentIndex = -1;
  for (var i = 0; i < categories.length; i++) {
    if (categories[i].no == select.value) {
      currentIndex = i;
      break;
    }
  }
  
  // ボタンの有効/無効を設定
  if (currentIndex === -1) {
    // カテゴリが見つからない場合
    prevButton.disabled = true;
    nextButton.disabled = true;
  } else {
    // 最初のカテゴリの場合
    prevButton.disabled = (currentIndex === 0);
    // 最後のカテゴリの場合
    nextButton.disabled = (currentIndex === categories.length - 1);
  }
}

// 選択数の表示を更新
function updateSelectionCount() {
  var selectionCount = document.getElementById('selectionCount');
  if (!selectionCount) return;
  
  var totalCount = currentCategoryData.length;
  var selectedCount = selectedQuestionIndices.length;
  
  if (selectedCount === 0) {
    // 未選択時は全問表示
    selectionCount.textContent = '全' + totalCount + '問';
    selectionCount.style.display = 'inline';
  } else {
    // 選択された問題のNoを取得して表示
    var selectedNos = [];
    selectedQuestionIndices.sort(function(a, b) { return a - b; }); // インデックスをソート
    selectedQuestionIndices.forEach(function(index) {
      if (index >= 0 && index < currentCategoryData.length) {
        var no = currentCategoryData[index].no;
        if (no) {
          selectedNos.push(no);
        }
      }
    });
    selectionCount.textContent = '全' + selectedCount + '問(' + selectedNos.join(',') + ')';
    selectionCount.style.display = 'inline';
  }
  
  // クリアボタンの有効/無効を更新
  updateClearButton();
}

// リスト表示をリセット
function resetListDisplay() {
  var listMessage = document.getElementById('listMessage');
  var listContainer = document.getElementById('listContainer');
  var startButton = document.getElementById('startButton');
  var selectionCount = document.getElementById('selectionCount');
  
  if (listMessage) listMessage.style.display = 'block';
  if (listContainer) listContainer.style.display = 'none';
  if (startButton) startButton.style.display = 'none';
  if (selectionCount) selectionCount.style.display = 'none';
  
  // ボタンを非表示
  hideListNavButtons();
  // ボタンの状態をリセット
  updateListNavButtons();
}

// エラーを表示
// エラーメッセージの配列（複数エラーを管理）
var errorMessages = [];

function showError(message) {
  // 既存のエラーメッセージコンテナを取得または作成
  var container = document.querySelector('.container');
  if (!container) return;
  
  var errorContainer = document.getElementById('errorContainer');
  if (!errorContainer) {
    errorContainer = document.createElement('div');
    errorContainer.id = 'errorContainer';
    container.insertBefore(errorContainer, container.firstChild);
  }
  
  // エラーメッセージを配列に追加
  errorMessages.push(message);
  
  // エラーメッセージを再描画
  renderErrorMessages();
}

// エラーメッセージを描画
function renderErrorMessages() {
  var errorContainer = document.getElementById('errorContainer');
  if (!errorContainer) return;
  
  // 既存のエラーメッセージを削除
  errorContainer.innerHTML = '';
  
  if (errorMessages.length === 0) {
    errorContainer.remove();
    return;
  }
  
  // エラーメッセージのdiv要素を作成
  var errorDiv = document.createElement('div');
  errorDiv.className = 'error-message';
  
  // 複数エラーの場合は箇条書きで表示
  if (errorMessages.length === 1) {
    errorDiv.textContent = errorMessages[0];
  } else {
    var ul = document.createElement('ul');
    errorMessages.forEach(function(msg) {
      var li = document.createElement('li');
      li.textContent = msg;
      ul.appendChild(li);
    });
    errorDiv.appendChild(ul);
  }
  
  // 閉じるボタンを追加
  var closeButton = document.createElement('button');
  closeButton.className = 'error-close-button';
  closeButton.textContent = '×';
  closeButton.type = 'button';
  closeButton.addEventListener('click', function() {
    clearErrorMessages();
  });
  errorDiv.appendChild(closeButton);
  
  errorContainer.appendChild(errorDiv);
}

// エラーメッセージをクリア
function clearErrorMessages() {
  errorMessages = [];
  var errorContainer = document.getElementById('errorContainer');
  if (errorContainer) {
    errorContainer.remove();
  }
}

// 学習開始
function startLearning() {
  if (currentCategoryData.length === 0) {
    return;
  }
  
  // 元のデータを保存
  originalCategoryData = currentCategoryData.slice();
  
  // 選択された問題のみを抽出（未選択時は全問）
  var filteredData = [];
  if (selectedQuestionIndices.length === 0) {
    // 未選択時は全問
    filteredData = currentCategoryData.slice();
  } else {
    // 選択された問題のみ（元の順序で）
    selectedQuestionIndices.sort(function(a, b) { return a - b; }); // インデックスをソート
    selectedQuestionIndices.forEach(function(index) {
      if (index >= 0 && index < currentCategoryData.length) {
        filteredData.push(currentCategoryData[index]);
      }
    });
  }
  
  // フィルタリングされたデータをcurrentCategoryDataに設定
  currentCategoryData = filteredData;
  
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
  
  // 再チャレンジ関連変数をリセット
  retryQuestionIndices = [];
  isInRetryMode = false;
  retryQuestionIndex = 0;
  completedQuestionIndices = [];
  isLearningCompleted = false;
  
  // 学習完了メッセージを非表示
  hideCompletionMessage();
  
  // 出題数表示を更新
  updateQuestionInfoDisplay();
  
  displayQuestion();
  
  // 最初の問題と次の問題をプリロード
  preloadAudioForCurrentAndNext();
  
  // トグルボタンの状態を引き継ぐ
  var toggleButton = document.getElementById('toggleButton');
  if (toggleButton) {
    if (isToggleActive) {
      toggleButton.classList.add('active');
    } else {
      toggleButton.classList.remove('active');
    }
  }
  
  // トグルボタンの位置を更新（screen2のタイトル位置に合わせる）
  requestAnimationFrame(function() {
    requestAnimationFrame(updateToggleButtonPosition);
  });
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
  updateQuestionInfoDisplay();
  
  // 質問文を表示（画像対応）
  var questionText = document.getElementById('questionText');
  if (questionText) {
    displayImageOrText(questionText, item.question || '');
  }
  
  // 上の黒いボックスを表示
  var answerButtonContainer = document.getElementById('answerButtonContainer');
  if (answerButtonContainer) answerButtonContainer.style.display = 'block';
  
  // ナビゲーションバーのAnswerボタンを有効化
  var navAnswerButton = document.getElementById('navAnswerButton');
  var navAnswerText = document.getElementById('navAnswerText');
  if (navAnswerButton) navAnswerButton.disabled = false;
  if (navAnswerText) navAnswerText.classList.add('blinking');
  
  // 回答テキストを非表示
  var answerTextDisplay = document.getElementById('answerTextDisplay');
  var noteSection = document.getElementById('noteSection');
  if (answerTextDisplay) answerTextDisplay.style.display = 'none';
  if (noteSection) noteSection.style.display = 'none';
  isAnswerShown = false;
  
  // 再生ボタンを無効化（問題表示時）
  var playButton = document.getElementById('playButton');
  if (playButton) {
    playButton.disabled = true;
  }
  
  // ストップウォッチをリセットして開始
  resetStopwatch();
  startStopwatch();
  
  // ナビゲーションボタンを無効化（Answerボタンが押されるまで）
  var nextButton = document.getElementById('nextButton');
  if (nextButton) nextButton.disabled = true;
  
  // プラスボタンを無効化（出題中）
  var plusButton = document.getElementById('plusButton');
  if (plusButton) plusButton.disabled = true;
  
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
  var navAnswerStopwatch = document.getElementById('navAnswerStopwatch');
  if (navAnswerStopwatch) {
    navAnswerStopwatch.textContent = '00:00:00';
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
  
  var timeText = String(minutes).padStart(2, '0') + ':' +
                 String(seconds).padStart(2, '0') + ':' +
                 String(milliseconds).padStart(2, '0');
  
  // ナビゲーションバーのAnswerボタン内のストップウォッチを更新
  var navAnswerStopwatch = document.getElementById('navAnswerStopwatch');
  if (navAnswerStopwatch) {
    navAnswerStopwatch.textContent = timeText;
  }
}

// 答えを表示
function showAnswer() {
  if (isAnswerShown) return;
  
  stopStopwatch();
  
  var item = currentCategoryData[currentQuestionIndex];
  
  // 上の黒いボックスを非表示
  var answerButtonContainer = document.getElementById('answerButtonContainer');
  if (answerButtonContainer) answerButtonContainer.style.display = 'none';
  
  // ナビゲーションバーのAnswerボタンを無効化
  var navAnswerButton = document.getElementById('navAnswerButton');
  var navAnswerText = document.getElementById('navAnswerText');
  if (navAnswerButton) navAnswerButton.disabled = true;
  if (navAnswerText) navAnswerText.classList.remove('blinking');
  
  // 回答文を表示（画像対応）
  var answerTextDisplay = document.getElementById('answerTextDisplay');
  if (answerTextDisplay) {
    displayImageOrText(answerTextDisplay, item.answer || '');
    answerTextDisplay.classList.remove('answer-hidden');
    answerTextDisplay.style.display = 'block';
  }
  
  // 再生ボタンの制御（回答が画像URLの場合は無効化）
  var playButton = document.getElementById('playButton');
  if (playButton) {
    if (isImageUrl(item.answer || '')) {
      // 画像URLの場合は無効化
      playButton.disabled = true;
    } else {
      // テキストの場合は有効化
      playButton.disabled = false;
    }
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
  
  // プラスボタンを有効化（回答表示中、学習完了でない場合）
  updatePlusButton();
  
  // トグルボタンがONの場合、自動再生
  if (isToggleActive) {
    playAnswer();
  }
}

/**
 * 画像URLかどうかを判定
 * @param {string} url - 判定する文字列
 * @returns {boolean} 画像URLの場合true
 */
function isImageUrl(url) {
  if (!url || typeof url !== 'string') return false;
  var trimmed = url.trim();
  return trimmed.startsWith('http://') || trimmed.startsWith('https://');
}

/**
 * Google Driveの共有リンクURLを直接表示用URLに変換
 * @param {string} url - Google Driveの共有リンクURL
 * @returns {string} 変換後のURL
 */
function convertGoogleDriveUrl(url) {
  if (!url || typeof url !== 'string') return url;
  
  // Google Driveの共有リンク形式を検出
  // 例: https://drive.google.com/file/d/FILE_ID/view?usp=drive_link
  var match = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    var fileId = match[1];
    // サムネイル形式を使用（広告ブロッカーにブロックされにくい）
    // sz=w1000 で最大幅1000pxの画像を取得（必要に応じて調整可能）
    return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';
  }
  
  // 変換不要の場合はそのまま返す
  return url;
}

/**
 * テキストまたは画像を表示
 * @param {HTMLElement} element - 表示先の要素
 * @param {string} content - 表示する内容（テキストまたは画像URL）
 */
function displayImageOrText(element, content) {
  if (!element || !content) {
    if (element) element.innerHTML = '';
    return;
  }
  
  var trimmedContent = content.trim();
  
  if (isImageUrl(trimmedContent)) {
    // 画像URLの場合
    var imageUrl = convertGoogleDriveUrl(trimmedContent);
    var img = document.createElement('img');
    img.src = imageUrl;
    img.className = 'content-image';
    img.alt = '画像';
    img.style.maxWidth = '100%';
    img.style.height = 'auto';
    img.style.display = 'block';
    img.style.margin = '0 auto';
    img.style.cursor = 'pointer';
    
    // クリックで拡大表示
    img.addEventListener('click', function() {
      showImageModal(imageUrl);
    });
    
    // エラーハンドリング
    img.addEventListener('error', function() {
      element.innerHTML = '<span class="image-error">画像を読み込めませんでした</span>';
    });
    
    element.innerHTML = '';
    element.appendChild(img);
  } else {
    // テキストの場合
    element.textContent = trimmedContent;
  }
}

/**
 * 画像を拡大表示（モーダル）
 * @param {string} imageUrl - 画像URL
 */
function showImageModal(imageUrl) {
  var overlay = document.getElementById('imageModalOverlay');
  var img = document.getElementById('imageModalImage');
  var closeButton = document.getElementById('imageModalCloseButton');
  
  if (!overlay || !img) return;
  
  img.src = imageUrl;
  overlay.style.display = 'flex';
  
  // 閉じるボタンのイベント
  if (closeButton) {
    closeButton.onclick = function() {
      overlay.style.display = 'none';
    };
  }
  
  // オーバーレイクリックで閉じる
  overlay.onclick = function(e) {
    if (e.target === overlay) {
      overlay.style.display = 'none';
    }
  };
  
  // ESCキーで閉じる
  var escHandler = function(e) {
    if (e.key === 'Escape') {
      overlay.style.display = 'none';
      document.removeEventListener('keydown', escHandler);
    }
  };
  document.addEventListener('keydown', escHandler);
}

// 回答を読み上げ
function playAnswer() {
  var item = currentCategoryData[currentQuestionIndex];
  if (!item || !item.answer) return;
  
  // 画像URLの場合は音声読み上げをスキップ
  if (isImageUrl(item.answer)) {
    return;
  }
  
  // WebアプリURLが設定されていない場合はエラー
  if (!WEB_APP_URL || WEB_APP_URL === 'YOUR_WEB_APP_URL_HERE') {
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
  // テキストを正規化（キャッシュキーは正規化後のテキストで生成）
  var normalizedText = normalizeTextForTTS(text);
  
  // メモリキャッシュを確認
  if (audioCache[normalizedText]) {
    return audioCache[normalizedText];
  }
  
  // localStorageを確認
  try {
    var cacheKey = CACHE_PREFIX + hashText(normalizedText);
    var cachedData = localStorage.getItem(cacheKey);
    if (cachedData) {
      var audioData = JSON.parse(cachedData);
      // メモリキャッシュにも保存
      audioCache[normalizedText] = audioData;
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
  // テキストを正規化（キャッシュキーは正規化後のテキストで生成）
  var normalizedText = normalizeTextForTTS(text);
  
  var audioData = {
    audioContent: audioContent,
    timestamp: Date.now(),
    textHash: hashText(normalizedText)  // メモリキャッシュ削除時の照合用
  };
  
  // メモリキャッシュに保存
  audioCache[normalizedText] = audioData;
  
  // localStorageに保存（サイズ制限を考慮）
  try {
    var cacheKey = CACHE_PREFIX + hashText(normalizedText);
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
      var entryKey = entries[j].key;
      var entryValue = localStorage.getItem(entryKey);
      
      // localStorageから削除
      localStorage.removeItem(entryKey);
      
      // メモリキャッシュからも削除（該当するものがあれば）
      if (entryValue) {
        try {
          var entryData = JSON.parse(entryValue);
          var storedHash = entryData.textHash;
          
          if (storedHash) {
            // textHashが保存されている場合（新形式）：ハッシュ値で直接照合
            for (var text in audioCache) {
              if (hashText(text) === storedHash) {
                delete audioCache[text];
                break; // 一致したらループを抜ける（効率化）
              }
            }
          } else {
            // textHashが保存されていない場合（旧形式）：従来の方法で照合
            var hashFromKey = entryKey.replace(CACHE_PREFIX, '');
            for (var text in audioCache) {
              if (hashText(text) === hashFromKey) {
                delete audioCache[text];
                break; // 一致したらループを抜ける（効率化）
              }
            }
          }
        } catch (e) {
          // パースエラーは無視（従来の方法でフォールバック）
          var hashFromKey = entryKey.replace(CACHE_PREFIX, '');
          for (var text in audioCache) {
            if (hashText(text) === hashFromKey) {
              delete audioCache[text];
              break;
            }
          }
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
    // 既存のAudioがあれば停止
    if (currentAudio) {
      currentAudio.pause();
      currentAudio.currentTime = 0;
      currentAudio = null;
    }
    
    var audio = new Audio('data:audio/mp3;base64,' + audioData.audioContent);
    currentAudio = audio;
    
    // 再生開始時に再生ボタンを無効化
    audio.addEventListener('play', function() {
      var playButton = document.getElementById('playButton');
      if (playButton) {
        playButton.disabled = true;
      }
    });
    
    // 再生終了時に再生ボタンを有効化
    audio.addEventListener('ended', function() {
      var playButton = document.getElementById('playButton');
      if (playButton) {
        // 回答が画像URLの場合は無効化のまま
        var item = currentCategoryData[currentQuestionIndex];
        if (item && !isImageUrl(item.answer || '')) {
          playButton.disabled = false;
        }
      }
      currentAudio = null;
    });
    
    // エラー時に再生ボタンを有効化
    audio.addEventListener('error', function() {
      var playButton = document.getElementById('playButton');
      if (playButton) {
        var item = currentCategoryData[currentQuestionIndex];
        if (item && !isImageUrl(item.answer || '')) {
          playButton.disabled = false;
        }
      }
      currentAudio = null;
    });
    
    audio.play().catch(function(error) {
      showError('音声の再生に失敗しました: ' + error.toString());
      // エラー時も再生ボタンを有効化
      var playButton = document.getElementById('playButton');
      if (playButton) {
        var item = currentCategoryData[currentQuestionIndex];
        if (item && !isImageUrl(item.answer || '')) {
          playButton.disabled = false;
        }
      }
      currentAudio = null;
    });
  } catch (error) {
    showError('音声の再生に失敗しました: ' + error.toString());
    currentAudio = null;
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
  
  // リクエストパラメータを準備
  var params = new URLSearchParams();
  params.append('text', text);
  params.append('email', userEmail); // TTS処理にもメール認証を追加
  params.append('referer', window.location.origin);
  
  // Google Apps Scriptにリクエストを送信
  fetch(WEB_APP_URL, {
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
      try {
        // 既存のAudioがあれば停止
        if (currentAudio) {
          currentAudio.pause();
          currentAudio.currentTime = 0;
          currentAudio = null;
        }
        
        var audio = new Audio('data:audio/mp3;base64,' + data.audioContent);
        currentAudio = audio;
        
        // 再生開始時に再生ボタンを無効化
        audio.addEventListener('play', function() {
          var playButton = document.getElementById('playButton');
          if (playButton) {
            playButton.disabled = true;
          }
        });
        
        // 再生終了時に再生ボタンを有効化
        audio.addEventListener('ended', function() {
          var playButton = document.getElementById('playButton');
          if (playButton) {
            // 回答が画像URLの場合は無効化のまま
            var item = currentCategoryData[currentQuestionIndex];
            if (item && !isImageUrl(item.answer || '')) {
              playButton.disabled = false;
            }
          }
          currentAudio = null;
        });
        
        // エラー時に再生ボタンを有効化
        audio.addEventListener('error', function() {
          var playButton = document.getElementById('playButton');
          if (playButton) {
            var item = currentCategoryData[currentQuestionIndex];
            if (item && !isImageUrl(item.answer || '')) {
              playButton.disabled = false;
            }
          }
          currentAudio = null;
        });
        
        audio.play().catch(function(error) {
          showError('音声の再生に失敗しました: ' + error.toString());
          // エラー時も再生ボタンを有効化
          var playButton = document.getElementById('playButton');
          if (playButton) {
            var item = currentCategoryData[currentQuestionIndex];
            if (item && !isImageUrl(item.answer || '')) {
              playButton.disabled = false;
            }
          }
          currentAudio = null;
        });
      } catch (error) {
        showError('音声の再生に失敗しました: ' + error.toString());
        currentAudio = null;
      }
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
  if (!WEB_APP_URL || WEB_APP_URL === 'YOUR_WEB_APP_URL_HERE') {
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
  if (!WEB_APP_URL || WEB_APP_URL === 'YOUR_WEB_APP_URL_HERE') {
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
    // userEmailが設定されていない場合はスキップ
    if (!userEmail) {
      return;
    }
    
    var params = new URLSearchParams();
    params.append('text', text);
    params.append('email', userEmail); // TTS処理にもメール認証を追加
    params.append('referer', window.location.origin);
    
    fetch(WEB_APP_URL, {
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
    });
  }, 100); // 少し遅延させて、メイン処理を優先
}

// 前の問題に戻る
function goToPreviousQuestion() {
  if (isInRetryMode) {
    // 再チャレンジモードの場合
    if (retryQuestionIndex > 0) {
      retryQuestionIndex--;
      currentQuestionIndex = retryQuestionIndices[retryQuestionIndex];
      displayQuestion();
      updateNavigationButtons();
    }
  } else {
    // 通常モード
    if (currentQuestionIndex > 0) {
      currentQuestionIndex--;
      displayQuestion();
      updateNavigationButtons();
    }
  }
}

// 次の問題に進む
function goToNextQuestion() {
  // 現在の問題を完了リストに追加（重複チェック）
  if (completedQuestionIndices.indexOf(currentQuestionIndex) === -1) {
    completedQuestionIndices.push(currentQuestionIndex);
  }
  
  if (isInRetryMode) {
    // 再チャレンジモードの場合、その問題をリストから削除
    var indexInRetry = retryQuestionIndices.indexOf(currentQuestionIndex);
    if (indexInRetry !== -1) {
      retryQuestionIndices.splice(indexInRetry, 1);
      // 削除後、現在のインデックスを調整
      if (retryQuestionIndex > indexInRetry) {
        retryQuestionIndex--;
      }
    }
    // 次の再チャレンジ問題があるか確認
    if (retryQuestionIndices.length > 0) {
      // 次の再チャレンジ問題に進む
      if (retryQuestionIndex < retryQuestionIndices.length) {
        currentQuestionIndex = retryQuestionIndices[retryQuestionIndex];
        displayQuestion();
      } else {
        // retryQuestionIndexが範囲外だが、再チャレンジ問題が残っている場合は再度再チャレンジを開始
        startRetryQuestions();
      }
    } else {
      // 再チャレンジ問題が全て終わった場合
      isInRetryMode = false;
      retryQuestionIndex = 0;
      isLearningCompleted = true;
      // 出題数表示を更新（完了済みとして表示）
      updateQuestionInfoDisplay();
      // 学習完了メッセージを表示
      showCompletionMessage();
    }
  } else {
    // 通常モード
    if (currentQuestionIndex < currentCategoryData.length - 1) {
      currentQuestionIndex++;
      displayQuestion();
    } else {
      // 最後の問題の場合
      if (retryQuestionIndices.length > 0) {
        // 再チャレンジ問題があれば表示
        startRetryQuestions();
      } else {
        // 再チャレンジ問題がなければ学習完了
        isLearningCompleted = true;
        // 出題数表示を更新（完了済みとして表示）
        updateQuestionInfoDisplay();
        // 学習完了メッセージを表示
        showCompletionMessage();
      }
    }
  }
  updateNavigationButtons();
  updatePlusButton();
}

// プラスボタンクリック処理
function handlePlusButtonClick() {
  if (isAnswerShown) {
    if (isInRetryMode) {
      // 再チャレンジモードの場合
      // 現在の問題を再チャレンジリストに追加（重複チェック）
      if (retryQuestionIndices.indexOf(currentQuestionIndex) === -1) {
        retryQuestionIndices.push(currentQuestionIndex);
      }
      
      // 完了リストから削除（プラスボタンを押したら黒色通常に戻す）
      var completedIndex = completedQuestionIndices.indexOf(currentQuestionIndex);
      if (completedIndex !== -1) {
        completedQuestionIndices.splice(completedIndex, 1);
      }
      
      // 出題数表示を更新
      updateQuestionInfoDisplay();
      
      // 次の再チャレンジ問題に進む
      retryQuestionIndex++;
      if (retryQuestionIndex < retryQuestionIndices.length) {
        // 次の再チャレンジ問題がある場合
        currentQuestionIndex = retryQuestionIndices[retryQuestionIndex];
        displayQuestion();
        updateNavigationButtons();
      } else {
        // 最後の再チャレンジ問題の場合、再チャレンジリストに残っている問題があれば再度再チャレンジを開始
        if (retryQuestionIndices.length > 0) {
          startRetryQuestions();
        } else {
          // 再チャレンジ問題がなければ学習完了
          isLearningCompleted = true;
          updateNavigationButtons();
          // 学習完了メッセージを表示
          showCompletionMessage();
        }
      }
    } else {
      // 通常モードの場合
      // 現在の問題を再チャレンジリストに追加（重複チェック）
      if (retryQuestionIndices.indexOf(currentQuestionIndex) === -1) {
        retryQuestionIndices.push(currentQuestionIndex);
      }
      
      // 完了リストから削除（プラスボタンを押したら黒色通常に戻す）
      var completedIndex = completedQuestionIndices.indexOf(currentQuestionIndex);
      if (completedIndex !== -1) {
        completedQuestionIndices.splice(completedIndex, 1);
      }
      
      // 出題数表示を更新
      updateQuestionInfoDisplay();
      
      // 次の問題に進む
      if (currentQuestionIndex < currentCategoryData.length - 1) {
        // 最後の問題でない場合、次の問題に進む
        currentQuestionIndex++;
        displayQuestion();
        updateNavigationButtons();
      } else {
        // 最後の問題の場合、最初に戻って再チャレンジ問題を出題
        startRetryQuestions();
      }
    }
  }
}

// 出題数表示を更新する関数
function updateQuestionInfoDisplay() {
  var questionInfo = document.getElementById('questionInfo');
  if (!questionInfo) return;
  
  // 元の全問題データを使用（選択されなかった問題も表示するため）
  var totalQuestions = originalCategoryData.length > 0 ? originalCategoryData.length : currentCategoryData.length;
  var displayItems = [];
  
  // 現在の問題が元のデータのどのインデックスに対応するかを取得
  var originalCurrentIndex = -1;
  if (currentQuestionIndex >= 0 && currentQuestionIndex < currentCategoryData.length) {
    var currentItem = currentCategoryData[currentQuestionIndex];
    // 元のデータから同じ問題を検索
    for (var idx = 0; idx < originalCategoryData.length; idx++) {
      if (originalCategoryData[idx] === currentItem || 
          (originalCategoryData[idx].no === currentItem.no && 
           originalCategoryData[idx].question === currentItem.question)) {
        originalCurrentIndex = idx;
        break;
      }
    }
  }
  
  // 選択された問題のインデックスを元のデータのインデックスに変換
  var originalSelectedIndices = [];
  if (selectedQuestionIndices.length > 0) {
    originalSelectedIndices = selectedQuestionIndices.slice();
  } else {
    // 未選択時は全問が選択されている
    for (var j = 0; j < totalQuestions; j++) {
      originalSelectedIndices.push(j);
    }
  }
  
  // retryQuestionIndicesを元のデータのインデックスに変換
  var originalRetryIndices = [];
  if (retryQuestionIndices.length > 0 && originalCategoryData.length > 0) {
    retryQuestionIndices.forEach(function(filteredIndex) {
      if (filteredIndex >= 0 && filteredIndex < currentCategoryData.length) {
        var retryItem = currentCategoryData[filteredIndex];
        // 元のデータから同じ問題を検索
        for (var retryIdx = 0; retryIdx < originalCategoryData.length; retryIdx++) {
          if (originalCategoryData[retryIdx] === retryItem || 
              (originalCategoryData[retryIdx].no === retryItem.no && 
               originalCategoryData[retryIdx].question === retryItem.question)) {
            originalRetryIndices.push(retryIdx);
            break;
          }
        }
      }
    });
  }
  
  // completedQuestionIndicesを元のデータのインデックスに変換
  var originalCompletedIndices = [];
  if (completedQuestionIndices.length > 0 && originalCategoryData.length > 0) {
    completedQuestionIndices.forEach(function(filteredIndex) {
      if (filteredIndex >= 0 && filteredIndex < currentCategoryData.length) {
        var completedItem = currentCategoryData[filteredIndex];
        // 元のデータから同じ問題を検索
        for (var completedIdx = 0; completedIdx < originalCategoryData.length; completedIdx++) {
          if (originalCategoryData[completedIdx] === completedItem || 
              (originalCategoryData[completedIdx].no === completedItem.no && 
               originalCategoryData[completedIdx].question === completedItem.question)) {
            originalCompletedIndices.push(completedIdx);
            break;
          }
        }
      }
    });
  }
  
  for (var i = 0; i < totalQuestions; i++) {
    var questionNum = i + 1;
    var isCurrent = (i === originalCurrentIndex);
    var isCompleted = (originalCompletedIndices.indexOf(i) !== -1);
    var isRetry = (originalRetryIndices.indexOf(i) !== -1);
    var isSelected = (originalSelectedIndices.indexOf(i) !== -1);
    
    if (!isSelected) {
      // 選択されなかった問題：グレー色
      displayItems.push('<span style="color: #999;">' + questionNum + '</span>');
    } else if (isCompleted) {
      // 完了済み（＞ボタンを押した）：灰色通常（最優先）
      displayItems.push('<span style="color: #999;">' + questionNum + '</span>');
    } else if (isCurrent && isRetry) {
      // 再チャレンジ問題を出題中：赤色太字
      displayItems.push('<strong style="color: #f00;">' + questionNum + '</strong>');
    } else if (isCurrent) {
      // 現在出題中：黒色太字
      displayItems.push('<strong style="color: #000;">' + questionNum + '</strong>');
    } else if (isRetry) {
      // 再チャレンジ対象（プラスボタンを押した）：赤色通常
      displayItems.push('<span style="color: #f00;">' + questionNum + '</span>');
    } else {
      // 未出題：黒色通常
      displayItems.push('<span style="color: #000;">' + questionNum + '</span>');
    }
  }
  
  questionInfo.innerHTML = displayItems.join(',');
}

// 再チャレンジ問題開始関数
function startRetryQuestions() {
  if (retryQuestionIndices.length > 0) {
    isInRetryMode = true;
    isLearningCompleted = false; // 再チャレンジ開始時は学習完了フラグをリセット
    retryQuestionIndex = 0;
    currentQuestionIndex = retryQuestionIndices[0];
    displayQuestion();
    updateNavigationButtons();
    updatePlusButton();
    // 学習完了メッセージを非表示
    hideCompletionMessage();
  } else {
    // 再チャレンジ問題がない場合は学習完了
    isLearningCompleted = true;
    updateNavigationButtons();
    updatePlusButton();
    // 学習完了メッセージを表示
    showCompletionMessage();
  }
}

// ナビゲーションボタンの状態を更新
function updateNavigationButtons() {
  var nextButton = document.getElementById('nextButton');
  
  // 回答表示中（isAnswerShown === true）の場合は、isLearningCompletedに関係なくボタンを有効化
  if (isAnswerShown && !isLearningCompleted) {
    if (isInRetryMode) {
      // 再チャレンジモードの場合
      // 最後の再チャレンジ問題でも、回答表示中は次へボタンを有効にする
      if (nextButton) nextButton.disabled = false;
    } else {
      // 通常モード
      if (nextButton) {
        if (currentQuestionIndex === currentCategoryData.length - 1) {
          // 最後の問題の場合、回答表示中は常に有効（再チャレンジ問題の有無に関係なく）
          nextButton.disabled = false;
        } else {
          nextButton.disabled = false;
        }
      }
    }
  } else if (isLearningCompleted) {
    // 学習完了の場合、次へボタンを無効化
    if (nextButton) nextButton.disabled = true;
  } else if (isInRetryMode) {
    // 再チャレンジモードの場合（回答表示前）
    if (nextButton) nextButton.disabled = true;
  } else {
    // 通常モード（回答表示前）
    if (nextButton) nextButton.disabled = true;
  }
}

// プラスボタンの状態を更新
function updatePlusButton() {
  var plusButton = document.getElementById('plusButton');
  if (plusButton) {
    // 回答表示中で学習完了でない場合は有効、それ以外は無効
    plusButton.disabled = !isAnswerShown || isLearningCompleted;
  }
}

// 学習完了メッセージを表示
function showCompletionMessage() {
  var completionSection = document.getElementById('completionMessageSection');
  var completionMessageText = document.querySelector('#completionMessage .completion-message-text');
  var completionMessageIcon = document.getElementById('completionMessageIcon');
  
  if (completionSection && completionMessageText && completionMessageIcon) {
    // メッセージをランダムに選択
    var randomMessageIndex = Math.floor(Math.random() * COMPLETION_MESSAGES.length);
    var message = COMPLETION_MESSAGES[randomMessageIndex];
    
    // アイコン画像をランダムに選択
    var randomImageIndex = Math.floor(Math.random() * COMPLETION_MESSAGE_IMAGES.length);
    var imageFileName = COMPLETION_MESSAGE_IMAGES[randomImageIndex];
    
    completionMessageText.textContent = message;
    completionMessageIcon.src = 'img/msg/' + imageFileName;
    completionSection.style.display = 'block';
  }
}

// 学習完了メッセージを非表示
function hideCompletionMessage() {
  var completionSection = document.getElementById('completionMessageSection');
  if (completionSection) {
    completionSection.style.display = 'none';
  }
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
  
  // トグルボタンの位置を更新（screen1に戻った時）
  requestAnimationFrame(function() {
    requestAnimationFrame(updateToggleButtonPosition);
  });
  
  // コンテナのパディングを元に戻す
  var container = document.querySelector('.container');
  if (container) container.classList.remove('learning-mode');
  
  // 選択状態をリセット
  selectedQuestionIndices = [];
  originalCategoryData = [];
  
  // 学習完了メッセージを非表示
  hideCompletionMessage();
  
  // 学習時間はリセットしない（継続）
}

// モーダルを表示
function showModal(item, index) {
  if (!item) return;
  
  // モーダル内の現在のインデックスを保存
  if (typeof index !== 'undefined') {
    modalCurrentIndex = index;
  } else {
    // インデックスが指定されていない場合は、itemから検索
    modalCurrentIndex = currentCategoryData.findIndex(function(data) {
      return data.id === item.id || (data.no === item.no && data.question === item.question);
    });
    if (modalCurrentIndex === -1) {
      modalCurrentIndex = 0;
    }
  }
  
  // モーダルの内容を更新
  updateModalContent(item);
  
  // ナビゲーションボタンの状態を更新
  updateModalNavigation();
  
  // 選択状態を更新
  updateModalSelection();
  
  // モーダルを表示
  var modalOverlay = document.getElementById('modalOverlay');
  if (modalOverlay) {
    modalOverlay.classList.add('active');
  }
}

// モーダルの内容を更新
function updateModalContent(item) {
  // A_Titleをラベルに設定
  if (item.a_title) {
    var modalAnswerLabel = document.getElementById('modalAnswerLabel');
    if (modalAnswerLabel) {
      modalAnswerLabel.textContent = item.a_title;
    }
  }
  
  // Q_Titleをラベルに設定
  if (item.q_title) {
    var modalQuestionLabel = document.getElementById('modalQuestionLabel');
    if (modalQuestionLabel) {
      modalQuestionLabel.textContent = item.q_title;
    }
  }
  
  // 質問文を表示（画像対応）
  var questionText = document.getElementById('modalQuestionText');
  if (questionText) {
    displayImageOrText(questionText, item.question || '');
  }
  
  // 回答文を表示（画像対応）
  var answerText = document.getElementById('modalAnswerText');
  if (answerText) {
    displayImageOrText(answerText, item.answer || '');
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
}

// モーダル内のナビゲーションを更新
function updateModalNavigation() {
  var totalCount = currentCategoryData.length;
  var currentNo = modalCurrentIndex + 1;
  
  // 現在No/全No数を更新
  var navInfo = document.getElementById('modalNavInfo');
  if (navInfo) {
    navInfo.textContent = currentNo + '/' + totalCount;
  }
  
  // 前へボタンの状態を更新
  var prevButton = document.getElementById('modalPrevButton');
  if (prevButton) {
    if (modalCurrentIndex === 0) {
      prevButton.disabled = true;
    } else {
      prevButton.disabled = false;
    }
  }
  
  // 次へボタンの状態を更新
  var nextButton = document.getElementById('modalNextButton');
  if (nextButton) {
    if (modalCurrentIndex === totalCount - 1) {
      nextButton.disabled = true;
    } else {
      nextButton.disabled = false;
    }
  }
}

// モーダルを閉じる
function closeModal() {
  var modalOverlay = document.getElementById('modalOverlay');
  if (modalOverlay) {
    modalOverlay.classList.remove('active');
  }
}

// モーダル内の選択状態を更新
function updateModalSelection() {
  var selectButton = document.getElementById('modalSelectButton');
  if (!selectButton) return;
  
  var isSelected = selectedQuestionIndices.indexOf(modalCurrentIndex) !== -1;
  if (isSelected) {
    selectButton.classList.add('selected');
  } else {
    selectButton.classList.remove('selected');
  }
}

// モーダル内の選択/解除を実行
function handleModalSelection() {
  var index = modalCurrentIndex;
  var selectedIndex = selectedQuestionIndices.indexOf(index);
  
  if (selectedIndex === -1) {
    // 選択
    selectedQuestionIndices.push(index);
  } else {
    // 解除
    selectedQuestionIndices.splice(selectedIndex, 1);
  }
  
  // モーダル内の選択状態を更新
  updateModalSelection();
  
  // リスト側の選択状態も更新
  updateListSelection(index);
  
  // 選択数の表示を更新
  updateSelectionCount();
}

// リスト側の選択状態を更新
function updateListSelection(index) {
  var tableBody = document.getElementById('listTableBody');
  if (!tableBody) return;
  
  var rows = tableBody.querySelectorAll('tr');
  if (index >= 0 && index < rows.length) {
    var row = rows[index];
    var noCell = row.querySelector('td:first-child');
    var isSelected = selectedQuestionIndices.indexOf(index) !== -1;
    
    if (isSelected) {
      row.classList.add('selected-row');
      if (noCell) noCell.classList.add('selected-no');
    } else {
      row.classList.remove('selected-row');
      if (noCell) noCell.classList.remove('selected-no');
    }
  }
}

// 選択をクリア
function clearSelection() {
  // 選択状態をクリア
  selectedQuestionIndices = [];
  
  // 全行の選択状態を解除
  var tableBody = document.getElementById('listTableBody');
  if (tableBody) {
    var rows = tableBody.querySelectorAll('tr');
    rows.forEach(function(row) {
      var noCell = row.querySelector('td:first-child');
      if (noCell) {
        noCell.classList.remove('selected-no');
      }
      row.classList.remove('selected-row');
    });
  }
  
  // 選択数表示を更新（クリアボタンの状態も更新される）
  updateSelectionCount();
}

// クリアボタンの有効/無効を更新
function updateClearButton() {
  var clearButton = document.getElementById('clearSelectionButton');
  if (!clearButton) return;
  
  // 選択がない場合は無効化
  clearButton.disabled = selectedQuestionIndices.length === 0;
}

// トグルコンテナの位置を更新（タイトルより少し下に配置）
function updateToggleButtonPosition() {
  var toggleContainer = document.getElementById('toggleContainer');
  if (!toggleContainer) return;
  
  // 現在表示されている画面のタイトルを取得
  var screen1 = document.getElementById('screen1');
  var screen2 = document.getElementById('screen2');
  var title = null;
  
  if (screen1 && screen1.classList.contains('active')) {
    title = screen1.querySelector('.title');
  } else if (screen2 && screen2.classList.contains('active')) {
    title = screen2.querySelector('.title');
  } else {
    // どちらもactiveでない場合は、表示されているタイトルを取得
    title = document.querySelector('.title');
  }
  
  if (!title) return;
  
  // タイトルの位置を取得
  var titleRect = title.getBoundingClientRect();
  
  // タイトルの中央の高さにコンテナを配置し、少し下に下げる（オフセット-18px）
  // コンテナの高さを考慮して中央揃え
  var toggleTop = titleRect.top + (titleRect.height / 2) - 18;
  
  toggleContainer.style.top = toggleTop + 'px';
}

