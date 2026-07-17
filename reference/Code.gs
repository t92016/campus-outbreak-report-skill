/**
 * ============================================================
 * 【範本】校園集體症狀通報系統 - 參考實作範例（原型：新北市忠義國小）
 * ⚠️ 這是一份『可直接執行的參考範例』，不是純佔位符樣板。
 *    請依照 docs/建置手冊.md 步驟一收集到的貴校規格，客製化下列標記為 TODO 的地方。
 * 完整 Apps Script 程式碼（單一檔案版）
 * ============================================================
 *
 * 使用方式：
 * 1. 開啟試算表 → 擴充功能 → Apps Script
 * 2. 把編輯器裡原本的內容「全部刪除」，貼上這整份程式碼
 * 3. 儲存（Ctrl+S）
 * 4. 依照下方「要執行的函式」清單，依序執行（每次執行後看下面的分節說明）
 *
 * ★ createReportForm 這個函式您已經執行過、表單已經建立了，
 *   請不要再執行第二次（會建立出重複的表單）。
 *   保留在這裡只是完整紀錄，方便日後查閱表單題目設定。
 */


// ============================================================
// 【設定區】所有共用常數，之後如果要調整欄位/選項，只需要改這裡
// ============================================================

const SHEET_MAIN = '彙整總表';        // 表單回應會先進到這裡
const SHEET_DEID = '去識別化總覽';    // 給首頁動態表格讀取用
const GRADE_LIST = ['幼兒園', '1年級', '2年級', '3年級', '4年級', '5年級', '6年級']; // 各年級分流工作表名稱
const FORM_ID = 'REPLACE_WITH_YOUR_FORM_ID'; // ★請在執行 createReportForm 後，把 Logger 印出的表單 ID 貼在這裡

// 學校 logo（已去背、256x256 PNG，base64 內嵌，首頁橫幅左上角使用）
const SCHOOL_LOGO_BASE64 = ''; // ★請依「建置手冊」步驟，把貴校 logo 去背後轉成 base64 貼在這裡（留空則首頁會顯示預設校徽 emoji）

// ============================================================
// ★ 2026-07-17 強化：彙整總表欄位改成「依標題文字動態尋找」，不再假設固定欄位編號
// ============================================================
// 原因：以後如果學校端（校長/老師）自己在 Google 表單新增/修改題目，欄位位置可能會跟
// 現在不一樣（例如刪除某一題後又重新建立同名題目，Google 會把它當成全新欄位加到最後面）。
// 用「標題文字」動態尋找欄位，不管實際欄位跑到第幾欄，只要標題文字沒變，程式都找得到，
// 大幅降低系統被表單異動搞壞的風險。
//
// HEADER_NAMES：每個欄位「代號」(key) 對應到「彙整總表標題列」實際會出現的文字。
const HEADER_NAMES = {
  TIMESTAMP: '時間戳記',   // A 表單自動
  GRADE: '年級',           // 表單自動
  CLASS: '班級',           // 表單自動
  SEAT: '座號',            // 表單自動
  NAME: '姓名',            // 表單自動
  GENDER: '性別',          // 表單自動
  CONTACT: '家人稱謂及緊急聯絡電話', // 表單自動
  SYMPTOM: '身體症狀',      // 表單自動
  CASE_ID: '案件編號',      // 程式產生
  SORT_KEY: '排序值',       // 程式產生
  TRIAGE: '檢傷初判',       // 檢傷組填寫
  LOCATION: '目前所在位置', // 休息組填寫
  HOSPITAL: '就醫醫院',     // 送醫組填寫
  ESCORT: '護送教師',       // 送醫組填寫
  NOTE: '備註',            // 三組皆可填寫
  STATUS: '狀態',          // 程式自動計算
  CLOSED: '結案',          // 程式自動計算，是/否
  UPDATED_AT: '最後更新時間', // 程式自動記錄
  UPDATED_BY: '最後更新人員'  // 三組操作時記錄
};

// 這些是系統運作的「核心必要欄位」，標題列如果找不到這些文字，代表表單/試算表結構
// 被改壞了，程式會丟出清楚的錯誤訊息，而不是默默算錯或存錯資料。
const REQUIRED_HEADER_KEYS = [
  'GRADE', 'CLASS', 'SEAT', 'NAME',
  'CASE_ID', 'SORT_KEY', 'STATUS', 'CLOSED', 'UPDATED_AT', 'UPDATED_BY'
];

/**
 * 動態讀取「彙整總表」（或其他傳入的 sheet）第一列的標題文字，
 * 建立「欄位代號 → 實際欄位編號」的對照表（回傳的物件用法跟以前的 COL 常數一模一樣，
 * 例如 COL.GRADE、COL.CASE_ID，只是現在是「即時查出來的」，不是寫死的數字）。
 *
 * @param {Sheet} sheet 要查詢的工作表（通常是彙整總表）
 * @param {boolean} throwOnMissing 找不到核心必要欄位時，是否要丟出錯誤（預設 true）
 */
function getMainColMap_(sheet, throwOnMissing) {
  if (throwOnMissing === undefined) throwOnMissing = true;

  const lastCol = sheet.getLastColumn();
  const headerRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const textToCol = {};
  headerRow.forEach(function (text, idx) {
    const t = String(text || '').trim();
    if (t) textToCol[t] = idx + 1; // 1-based
  });

  const map = {};
  Object.keys(HEADER_NAMES).forEach(function (key) {
    const headerText = HEADER_NAMES[key];
    map[key] = textToCol[headerText] || null;
  });
  map.LAST_COL = lastCol; // 實際總欄數（含任何學校自行新增、我們沒有對應到的題目欄位）

  if (throwOnMissing) {
    const missing = REQUIRED_HEADER_KEYS.filter(function (key) { return !map[key]; });
    if (missing.length) {
      const missingText = missing.map(function (k) { return HEADER_NAMES[k]; }).join('、');
      throw new Error(
        '「' + sheet.getName() + '」找不到必要欄位標題：[' + missingText + ']。' +
        '可能是 Google 表單或試算表標題列被誤改/誤刪，請檢查標題列文字是否跟系統預期的一致，' +
        '或聯絡系統維護者協助排查（不要自行嘗試修改欄位順序）。'
      );
    }
  }

  return map;
}

// 選項清單（需與表單/三組系統的下拉選單一致）
const OPTIONS = {
  TRIAGE: ['休息觀察區', '送醫區'],
  LOCATION: [
    '休息觀察區(1F國際文教中心)',
    '送醫區(1F楊陳包藝文中心)',
    '已回教室',
    '回檢傷組重新判斷',
    '家長接回'
  ],
  HOSPITAL: [
    '家長接回',
    '蘆洲衛生所',
    '新北市立聯合醫院(三重)',
    '三重衛生所',
    '三重中興醫院',
    '三重宏仁醫院',
    '淡水馬偕醫院',
    '士林新光醫院',
    '中山馬偕醫院',
    '大同中興醫院',
    '中正臺大醫院',
    '北投振興醫院',
    '北投榮民醫院',
    '新莊新仁醫院',
    '新莊臺北醫院',
    '新莊新泰醫院',
    '林口長庚醫院',
    '松山長庚醫院',
    '已回教室'
  ]
};

// 各年級分流工作表的欄位順序（必須跟 buildGradeSheetHeaders_ 的順序完全一致！）
// 用來讓「反向同步」知道年級分頁的第 N 欄，對應彙整總表的哪一個欄位（COL 的哪個 key）
const GRADE_SHEET_FIELDS = [
  'CASE_ID', 'CLASS', 'SEAT', 'NAME', 'GENDER', 'CONTACT', 'SYMPTOM',
  'TRIAGE', 'LOCATION', 'HOSPITAL', 'ESCORT', 'NOTE',
  'STATUS', 'CLOSED', 'UPDATED_AT', 'UPDATED_BY', 'SORT_KEY'
];

// 年級分頁裡，允許「反向寫回」彙整總表的欄位（其餘欄位是系統自動計算，編輯了也會被下次同步覆蓋）
const GRADE_EDITABLE_FIELDS = [
  'CLASS', 'SEAT', 'NAME', 'GENDER', 'CONTACT', 'SYMPTOM',
  'TRIAGE', 'LOCATION', 'HOSPITAL', 'ESCORT', 'NOTE'
];

// ---------- 三組通報網頁設定 ----------
// propKey 對應到 Apps Script「指令碼屬性」(Project Settings > Script Properties) 裡的密碼欄位名稱，
// 請自行到 Apps Script 編輯器 → 左側齒輪「專案設定」→「指令碼屬性」→ 新增這三筆屬性並填入密碼。
const GROUP_CONFIG = {
  triage: {
    label: '檢傷組', propKey: 'PWD_TRIAGE',
    themeColor: '#e8a598', themeColorLight: '#f7e3df',
    updateFieldLabel: '檢傷初判', options: OPTIONS.TRIAGE
  },
  rest: {
    label: '休息組', propKey: 'PWD_REST',
    themeColor: '#9db3ab', themeColorLight: '#e6ecea',
    updateFieldLabel: '目前所在位置', options: OPTIONS.LOCATION
  },
  hospital: {
    label: '送醫組', propKey: 'PWD_HOSPITAL',
    themeColor: '#b8ab6e', themeColorLight: '#efeadd',
    updateFieldLabel: '就醫醫院', options: OPTIONS.HOSPITAL
  }
};

// 管理者密碼屬性名稱（請至 Apps Script「指令碼屬性」自行新增 PWD_ADMIN 並設定密碼）
const ADMIN_PROP_KEY = 'PWD_ADMIN';

// 年級 → 數字（給排序值計算用）
function gradeToNumber_(gradeText) {
  const map = {
    '幼兒園': 0, '1年級': 1, '2年級': 2, '3年級': 3,
    '4年級': 4, '5年級': 5, '6年級': 6
  };
  return (gradeText in map) ? map[gradeText] : 99;
}


// ============================================================
// 【步驟 0：已完成，請勿重複執行】建立 Google 表單
// ============================================================
function createReportForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = FormApp.create('○○國小校園出現集體腸胃不適症狀通報表'); // TODO: 換成貴校表單標題（依開發規格步驟1收集）

  form.setDescription(
    '1.全校師生若於午餐後發生疑似食物中毒情形，請填寫表單通報個案資料，並即時告知學務處(分機137)。\n' +
    '2.每個個案請個別填寫一張表單。\n' +
    '3.導師請填1~6題後提交，其餘內容由相關各組人員填寫。'
  );
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);
  form.setProgressBar(false);
  form.setShowLinkToRespondAgain(true);

  form.addListItem().setTitle('年級')
    .setChoiceValues(['1年級', '2年級', '3年級', '4年級', '5年級', '6年級', '幼兒園'])
    .setRequired(true);

  const classChoices = [];
  for (let i = 1; i <= 17; i++) classChoices.push(String(i));
  form.addListItem().setTitle('班級').setChoiceValues(classChoices).setRequired(true);

  const seatChoices = [];
  for (let i = 1; i <= 34; i++) seatChoices.push(String(i));
  form.addListItem().setTitle('座號').setChoiceValues(seatChoices).setRequired(true);

  form.addTextItem().setTitle('姓名').setRequired(true);

  form.addTextItem().setTitle('家人稱謂及緊急聯絡電話')
    .setHelpText('例如：媽媽 0911333666').setRequired(true);

  form.addCheckboxItem().setTitle('身體症狀')
    .setChoiceValues(['腹瀉', '嘔吐', '噁心', '腹痛', '發燒', '頭痛', '紅疹'])
    .showOtherOption(true).setRequired(true);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('表單網址：' + form.getPublishedUrl());
  Logger.log('表單編輯網址：' + form.getEditUrl());
  Logger.log('表單 ID：' + form.getId());
}


// ============================================================
// 【步驟 1：請執行】在表單中新增「性別」題目
// ============================================================
function addGenderQuestion() {
  const form = FormApp.openById(FORM_ID);

  form.addListItem()
    .setTitle('性別')
    .setChoiceValues(['男', '女'])
    .setRequired(true);

  const items = form.getItems();
  const genderItem = items[items.length - 1];
  form.moveItem(genderItem, 4); // 移動到「姓名」之後

  Logger.log('已新增「性別」題目，並移動到姓名之後。');
}


// ============================================================
// 【步驟 2：請執行】初始化試算表結構
// ============================================================
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 自動確認/修正：表單的回覆目的地必須是「目前這份試算表」
  ensureFormLinkedToThisSpreadsheet_(ss);

  const mainSheet = findOrRenameFormResponseSheet_(ss);

  // ★ 動態決定「系統自動欄位」要放在第幾欄開始：
  //   如果之前已經設定過（標題列已經有「案件編號」），就沿用原本的位置，不會搬動；
  //   如果是第一次設定，就接在目前表單既有欄位的最後面。
  const existingMap = getMainColMap_(mainSheet, false);
  const startCol = existingMap.CASE_ID || (mainSheet.getLastColumn() + 1);

  const headers = [
    '案件編號', '排序值', '檢傷初判', '目前所在位置', '就醫醫院',
    '護送教師', '備註', '狀態', '結案', '最後更新時間', '最後更新人員'
  ];
  mainSheet.getRange(1, startCol, 1, headers.length).setValues([headers]);
  mainSheet.getRange(1, 1, 1, startCol + headers.length - 1)
    .setFontWeight('bold').setBackground('#c0645c').setFontColor('#ffffff');
  mainSheet.setFrozenRows(1);

  GRADE_LIST.forEach(function (gradeName) {
    getOrCreateSheet_(ss, gradeName, buildGradeSheetHeaders_());
  });

  const deidSheet = getOrCreateSheet_(ss, SHEET_DEID, [
    '班級-座號', '姓名', '性別', '狀態', '休息觀察', '送醫醫院', '家長帶離醫院', '出院時間', '排序值'
  ]);
  protectSheetWithWarning_(deidSheet, '此工作表為「去識別化總覽」，資料由系統自動產生（姓名已去識別化），請勿手動編輯，修改內容會在下次同步時被覆蓋，且可能造成真實姓名對應錯誤。');

  // 對「彙整總表」的系統自動欄位加上編輯警告：依剛剛實際寫入的位置動態計算範圍
  const COL = getMainColMap_(mainSheet);
  protectRangeWithWarning_(
    mainSheet.getRange(1, COL.CASE_ID, mainSheet.getMaxRows(), 2), // 案件編號、排序值（連續兩欄）
    '「案件編號」「排序值」為系統自動產生/計算欄位，請勿手動修改，修改後系統會依規則自動同步覆蓋。'
  );
  protectRangeWithWarning_(
    mainSheet.getRange(1, COL.STATUS, mainSheet.getMaxRows(), 4), // 狀態、結案、最後更新時間、最後更新人員（連續四欄）
    '「狀態」「結案」「最後更新時間」「最後更新人員」為系統自動計算欄位，請勿手動修改，修改後系統會自動重新計算覆蓋。'
  );

  // 對各年級分頁的系統自動欄位加上編輯警告：案件編號(A) 單獨一欄；狀態~排序值(M~Q) 連續欄位
  GRADE_LIST.forEach(function (gradeName) {
    const gradeSheet = ss.getSheetByName(gradeName);
    protectRangeWithWarning_(
      gradeSheet.getRange(1, 1, gradeSheet.getMaxRows(), 1), // A 案件編號
      '「案件編號」為系統自動產生欄位，請勿手動修改，修改後系統會自動同步覆蓋回正確值。'
    );
    protectRangeWithWarning_(
      gradeSheet.getRange(1, 13, gradeSheet.getMaxRows(), 5), // M:Q 狀態、結案、最後更新時間、最後更新人員、排序值
      '「狀態」「結案」「最後更新時間」「最後更新人員」「排序值」為系統自動計算欄位，請勿手動修改，修改後系統會自動重新計算覆蓋。'
    );
  });

  SpreadsheetApp.getUi().alert('初始化完成！請檢查「彙整總表」欄位與各年級/去識別化工作表是否已建立。');
}

/**
 * 幫「整張工作表」加上「編輯警告」保護：使用者編輯時會跳出警告訊息，
 * 但不會真的擋住編輯（避免我們自己的程式同步時被卡住），
 * 純粹提醒使用者「這裡是自動產生的表，不建議手動改」。
 */
function protectSheetWithWarning_(sheet, message) {
  removeExistingProtectionsWithDescription_(sheet, message);
  const protection = sheet.protect().setDescription(message);
  protection.setWarningOnly(true);
}

/**
 * 幫「特定範圍/欄位」加上「編輯警告」保護（同上，只是範圍縮小到指定欄位，
 * 而不是鎖住整張工作表，這樣同一張表裡其他允許編輯的欄位才不會一起被警告）。
 */
function protectRangeWithWarning_(range, message) {
  removeExistingProtectionsWithDescription_(range.getSheet(), message);
  const protection = range.protect().setDescription(message);
  protection.setWarningOnly(true);
}

/**
 * 避免重複執行 setupSpreadsheet 時疊加出多個相同描述的保護範圍。
 */
function removeExistingProtectionsWithDescription_(sheet, message) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET));
  protections.forEach(function (p) {
    if (p.getDescription() === message) p.remove();
  });
}

/**
 * 確認表單的回覆目的地是不是「目前這份試算表」，如果不是（或還沒設定），
 * 就自動重新連結，並稍微等待讓 Google 有時間建立回應工作表。
 */
function ensureFormLinkedToThisSpreadsheet_(ss) {
  const form = FormApp.openById(FORM_ID);
  const destId = form.getDestinationId(); // 目前連結的試算表 ID，若未連結則為 null

  if (destId !== ss.getId()) {
    Logger.log('表單目前的回覆目的地(' + destId + ')與本試算表(' + ss.getId() + ')不同，正在自動重新連結...');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    Utilities.sleep(2000); // 給 Google 一點時間建立/連結回應工作表
  }
}

function findOrRenameFormResponseSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_MAIN);
  if (sheet) return sheet;

  const sheets = ss.getSheets();
  const sheetNames = sheets.map(function (s) { return s.getName(); });

  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (name.indexOf('表單回應') !== -1 || name.indexOf('表單回覆') !== -1) {
      sheets[i].setName(SHEET_MAIN);
      return sheets[i];
    }
  }

  throw new Error(
    '找不到表單回應工作表。目前這份試算表裡的工作表有：[' + sheetNames.join(', ') + ']。' +
    '請確認您是在「表單要連結的那一份」試算表裡執行本程式碼，或聯絡開發者協助排查。'
  );
}

function getOrCreateSheet_(ss, name, headerRow) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  sheet.getRange(1, 1, 1, headerRow.length)
    .setFontWeight('bold').setBackground('#c0645c').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return sheet;
}

function buildGradeSheetHeaders_() {
  return [
    '案件編號', '班級', '座號', '姓名', '性別', '聯絡電話', '身體症狀',
    '檢傷初判', '目前所在位置', '就醫醫院', '護送教師', '備註',
    '狀態', '結案', '最後更新時間', '最後更新人員', '排序值'
  ];
}


// ============================================================
// 【步驟 3：請執行】安裝自動觸發器（表單送出時 ＋ 手動編輯彙整總表時）
// ============================================================
function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ['onFormSubmitHandler', 'onEditHandler', 'onChangeHandler'].forEach(function (fnName) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === fnName) {
        ScriptApp.deleteTrigger(t);
      }
    });
  });

  ScriptApp.newTrigger('onFormSubmitHandler')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // onChange：處理「整列刪除/插入」這類結構性變更（onEdit 不會觸發，需要另外監聽）
  ScriptApp.newTrigger('onChangeHandler')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  SpreadsheetApp.getUi().alert(
    '已安裝觸發器：\n' +
    '1) 表單送出時 → 自動處理新資料\n' +
    '2) 手動編輯「彙整總表」時 → 自動同步到各年級表與去識別化總覽\n' +
    '3) 試算表結構變更（例如整列刪除）時 → 自動同步\n\n' +
    '請注意：只有編輯「彙整總表」才會觸發同步。各年級分頁與去識別化總覽是自動產生的檢視表，' +
    '請不要直接在那些分頁手動修改，修改了也會在下次同步時被覆蓋。'
  );
}

/**
 * 【結構性變更時自動執行】不需要手動執行，是觸發器自動呼叫的。
 * 處理「整列刪除/插入」這類 onEdit 不會觸發的情況，確保刪除學生資料後
 * 各年級分頁、去識別化總覽（首頁）都會自動同步移除。
 */
function onChangeHandler(e) {
  syncDerivedSheets();
}

/**
 * 【手動編輯時自動執行】不需要手動執行，是觸發器自動呼叫的。
 *
 * - 編輯「彙整總表」 → 直接重新同步全部衍生工作表
 * - 編輯「各年級分頁」 → 依「案件編號」回頭找到彙整總表對應的那一列，
 *                        把可編輯欄位的異動寫回去，再重新同步全部工作表
 * - 編輯「去識別化總覽」 → 不做任何事（此表唯讀，姓名已遮蔽，不應寫回真實資料）
 *
 * 注意：Apps Script 用程式（例如 syncDerivedSheets 本身）寫入儲存格不會觸發 onEdit，
 *       只有「真人」在畫面上編輯儲存格才會觸發，所以不會造成無限迴圈。
 */
function onEditHandler(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === SHEET_MAIN) {
    syncDerivedSheets();
    return;
  }

  if (GRADE_LIST.indexOf(sheetName) !== -1) {
    propagateGradeEditToMain_(sheet, e.range);
    return;
  }

  // SHEET_DEID 或其他工作表：不處理
}

/**
 * 把「年級分頁」上的手動編輯，依案件編號寫回「彙整總表」對應的列，
 * 並更新最後更新時間/人員，最後重新整站同步。
 */
function propagateGradeEditToMain_(gradeSheet, editedRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);

  const startRow = editedRange.getRow();
  const numRows = editedRange.getNumRows();
  const startCol = editedRange.getColumn();
  const numCols = editedRange.getNumColumns();

  for (let i = 0; i < numRows; i++) {
    const gRow = startRow + i;
    if (gRow === 1) continue; // 跳過標題列

    const caseId = gradeSheet.getRange(gRow, 1).getValue(); // 第1欄固定是「案件編號」
    if (!caseId) continue; // 空白列，略過

    const mainRow = findMainRowByCaseId_(mainSheet, caseId, COL);
    if (!mainRow) continue; // 找不到對應資料（可能案件編號被手動改壞、或剛好還沒同步），略過

    let rowChanged = false;
    for (let j = 0; j < numCols; j++) {
      const gCol = startCol + j;
      const fieldKey = GRADE_SHEET_FIELDS[gCol - 1];
      if (!fieldKey) continue;
      if (GRADE_EDITABLE_FIELDS.indexOf(fieldKey) === -1) continue; // 非可編輯欄位，略過（系統自動計算，不寫回）

      const newValue = gradeSheet.getRange(gRow, gCol).getValue();
      const mainCol = COL[fieldKey];
      mainSheet.getRange(mainRow, mainCol).setValue(newValue);
      rowChanged = true;
    }

    if (rowChanged) {
      mainSheet.getRange(mainRow, COL.UPDATED_AT).setValue(new Date());
      mainSheet.getRange(mainRow, COL.UPDATED_BY).setValue('人工修改(' + gradeSheet.getName() + ')');
    }
  }

  // ★ 不論這次編輯有沒有改到「可編輯欄位」，都一律重新整站同步。
  //   這樣就算有人手動改到「案件編號/狀態/排序值」這類系統自動欄位，
  //   也會在同步時被自動修正回正確值，不會卡住、也不會一直維持錯誤狀態。
  syncDerivedSheets();
}

/**
 * 依「案件編號」在彙整總表裡找到對應的列號（找不到回傳 null）。
 * colMap 可選填：如果外面已經算好 COL（getMainColMap_ 的結果）可以直接傳進來，
 * 避免同一次操作裡重複讀取標題列；沒傳的話會自己重新算一次。
 */
function findMainRowByCaseId_(mainSheet, caseId, colMap) {
  const COL = colMap || getMainColMap_(mainSheet);
  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return null;
  const ids = mainSheet.getRange(2, COL.CASE_ID, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === caseId) return i + 2;
  }
  return null;
}


// ============================================================
// 【表單送出時自動執行】不需要手動執行，是觸發器自動呼叫的
// ============================================================
function onFormSubmitHandler(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(sheet);
  const row = e.range.getRow();

  const grade = sheet.getRange(row, COL.GRADE).getValue();
  const classNo = sheet.getRange(row, COL.CLASS).getValue();
  const seatNo = sheet.getRange(row, COL.SEAT).getValue();
  const now = new Date();

  const caseId = generateCaseId_(sheet, now, COL);
  const sortKey = gradeToNumber_(grade) * 10000 + Number(classNo) * 100 + Number(seatNo);

  sheet.getRange(row, COL.CASE_ID).setValue(caseId);
  sheet.getRange(row, COL.SORT_KEY).setValue(sortKey);
  sheet.getRange(row, COL.STATUS).setValue('檢傷中');
  sheet.getRange(row, COL.CLOSED).setValue('否');
  sheet.getRange(row, COL.UPDATED_AT).setValue(now);
  sheet.getRange(row, COL.UPDATED_BY).setValue('導師通報');

  syncDerivedSheets();
}


// ============================================================
// 【全站資料同步】表單送出、或三組更新資料後都會呼叫這個函式
// ============================================================
function syncDerivedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return;

  // 讀取寬度用「實際總欄數」（COL.LAST_COL），這樣就算學校自己在表單新增了
  // 我們沒有對應到的新題目（欄位會被 Google 加到最後面），也不會漏讀/讀錯既有欄位。
  const range = mainSheet.getRange(2, 1, lastRow - 1, COL.LAST_COL);
  const values = range.getValues();

  const updatedValues = values.map(function (r) {
    // 重新計算排序值（若有人手動修正年級/班級/座號，這裡會自動反映最新的分流與排序）
    const grade = r[COL.GRADE - 1];
    const classNo = r[COL.CLASS - 1];
    const seatNo = r[COL.SEAT - 1];
    if (grade && classNo !== '' && seatNo !== '') {
      r[COL.SORT_KEY - 1] = gradeToNumber_(grade) * 10000 + Number(classNo) * 100 + Number(seatNo);
    }

    const triage = r[COL.TRIAGE - 1];
    const location = r[COL.LOCATION - 1];
    const hospital = r[COL.HOSPITAL - 1];
    const status = computeStatus_(triage, location, hospital);
    r[COL.STATUS - 1] = status.text;
    r[COL.CLOSED - 1] = status.closed ? '是' : '否';
    return r;
  });

  // 依排序值排序（由小到大：幼兒園→1年級→...→6年級，同年級再依班級、座號）
  const sorted = updatedValues.slice().sort(function (a, b) {
    return (a[COL.SORT_KEY - 1] || 0) - (b[COL.SORT_KEY - 1] || 0);
  });

  // 把排序後的結果寫回「彙整總表」本身，讓總表也保持排序狀態
  range.setValues(sorted);

  GRADE_LIST.forEach(function (gradeName) {
    const gradeRows = sorted.filter(function (r) { return r[COL.GRADE - 1] === gradeName; });
    writeGradeSheet_(ss, gradeName, gradeRows, COL);
  });

  writeDeidSheet_(ss, sorted, COL);
}

function computeStatus_(triage, location, hospital) {
  // ★ 2026-07-17 修正：以前的邏輯要求「檢傷初判」一定要先填，才會去看「目前所在位置」
  //   跟「就醫醫院」的內容，但實務上（尤其是管理者後台手動建立/修改測試資料時）可能
  //   會出現「檢傷初判是空的，但休息觀察或送醫醫院已經有結果」這種不照順序的情況，
  //   導致狀態一直卡在「檢傷中」，跟休息觀察/送醫醫院欄位顯示的結果矛盾。
  //   現在改成優先看「實際已發生的結果」（休息觀察/送醫醫院的內容），
  //   不管檢傷初判是否已填寫，只要有更後續的紀錄，就以那個結果為準。

  // 最終結案結果：已回教室 / 家長接回，只要任一欄位出現這個結果就視為已結案
  if (hospital === '已回教室' || location === '已回教室') {
    return { text: '已回教室', closed: true };
  }
  if (hospital === '家長接回' || location === '家長接回') {
    return { text: '家長接回', closed: true };
  }

  // 就醫醫院已有記錄（不論檢傷初判是否已填），代表確實已送醫
  if (hospital) {
    return { text: '送醫中(已抵達' + hospital + ')', closed: false };
  }
  if (triage === '送醫區') {
    return { text: '送醫中', closed: false };
  }

  // 目前所在位置已有記錄
  if (location) {
    if (location.indexOf('休息觀察區') === 0) return { text: '休息', closed: false };
    if (location === '回檢傷組重新判斷') return { text: '檢傷中', closed: false };
    if (location.indexOf('送醫區') === 0) return { text: '送醫中', closed: false };
    return { text: '休息', closed: false };
  }
  if (triage === '休息觀察區') {
    return { text: '休息', closed: false };
  }

  return { text: '檢傷中', closed: false };
}

function generateCaseId_(sheet, dateObj, colMap) {
  const COL = colMap || getMainColMap_(sheet);
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');

  const lastRow = sheet.getLastRow();
  let countToday = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, COL.CASE_ID, lastRow - 1, 1).getValues();
    ids.forEach(function (r) {
      if (String(r[0]).indexOf(dateStr) === 0) countToday++;
    });
  }
  const seq = String(countToday + 1).padStart(3, '0');
  return dateStr + '-' + seq;
}

function deidentifyName_(name) {
  if (!name) return '';
  const s = String(name).trim();
  if (s.length <= 1) return s;
  if (s.length === 2) return s.charAt(0) + '○';
  const middle = '○'.repeat(s.length - 2);
  return s.charAt(0) + middle + s.charAt(s.length - 1);
}

function formatClassSeatCode_(grade, classNo, seatNo) {
  const g = gradeToNumber_(grade);
  const c = String(classNo).padStart(2, '0');
  const s = String(seatNo).padStart(2, '0');
  return String(g) + c + '-' + s;
}

function writeGradeSheet_(ss, gradeName, rows, COL) {
  const sheet = ss.getSheetByName(gradeName);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  if (rows.length === 0) return;

  const output = rows.map(function (r) {
    return [
      r[COL.CASE_ID - 1], r[COL.CLASS - 1], r[COL.SEAT - 1], r[COL.NAME - 1], r[COL.GENDER - 1],
      r[COL.CONTACT - 1], r[COL.SYMPTOM - 1], r[COL.TRIAGE - 1], r[COL.LOCATION - 1],
      r[COL.HOSPITAL - 1], r[COL.ESCORT - 1], r[COL.NOTE - 1], r[COL.STATUS - 1],
      r[COL.CLOSED - 1], r[COL.UPDATED_AT - 1], r[COL.UPDATED_BY - 1], r[COL.SORT_KEY - 1]
    ];
  });
  sheet.getRange(2, 1, output.length, output[0].length).setValues(output);
}

function writeDeidSheet_(ss, sortedRows, COL) {
  const sheet = ss.getSheetByName(SHEET_DEID);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  if (sortedRows.length === 0) return;

  const output = sortedRows.map(function (r) {
    const classSeatCode = formatClassSeatCode_(r[COL.GRADE - 1], r[COL.CLASS - 1], r[COL.SEAT - 1]);
    const deidName = deidentifyName_(r[COL.NAME - 1]);
    const status = r[COL.STATUS - 1];
    const location = r[COL.LOCATION - 1] || '';
    const hospital = r[COL.HOSPITAL - 1] || '';
    const parentPickup = (hospital === '家長接回' || location === '家長接回') ? '是' : '';
    const dischargeTime = (status === '已回教室' || status === '家長接回') ? r[COL.UPDATED_AT - 1] : '';

    return [
      classSeatCode, deidName, r[COL.GENDER - 1] || '', status,
      location, hospital, parentPickup, dischargeTime, r[COL.SORT_KEY - 1]
    ];
  });
  sheet.getRange(2, 1, output.length, output[0].length).setValues(output);
}


// ============================================================
// 【網頁應用程式】doGet 路由：首頁 / 檢傷組 / 休息組 / 送醫組
// ============================================================
// 部署方式：Apps Script 編輯器右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
// 「執行身分」選「我」，「誰可以存取」選「所有人」，部署後會得到一個網址，
// 網址後面加上 ?page=triage / ?page=rest / ?page=hospital 就能切換到各組頁面，
// 不加任何參數就是首頁動態總覽表格。
function doGet(e) {
  const page = ((e && e.parameter && e.parameter.page) || '').toLowerCase();

  if (page === 'admin') {
    return HtmlService.createHtmlOutput(buildAdminHtml_())
      .setTitle('管理者後台')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (GROUP_CONFIG[page]) {
    // ★ 手機版關鍵修正：一定要加上 viewport meta 標籤，
    //   否則手機瀏覽器會用桌面版寬度渲染再縮小，下拉選單/按鈕會小到無法點選。
    return HtmlService.createHtmlOutput(buildGroupPortalHtml_(page))
      .setTitle(GROUP_CONFIG[page].label + '通報')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutput(buildHomeHtml_())
    .setTitle('校園出現集體腸胃不適症狀事件通報') // TODO: 換成貴校系統標題
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// 【網頁應用程式後端函式】由前端 HTML 裡的 google.script.run 呼叫
// ============================================================

/**
 * 驗證組別密碼是否正確。
 * 密碼存在 Apps Script「指令碼屬性」，請至 Apps Script 編輯器
 * → 左側齒輪「專案設定」→「指令碼屬性」新增：
 *   PWD_TRIAGE（檢傷組）、PWD_REST（休息組）、PWD_HOSPITAL（送醫組）
 */
function verifyGroupPassword(group, pwd) {
  const cfg = GROUP_CONFIG[group];
  if (!cfg) return false;
  const stored = PropertiesService.getScriptProperties().getProperty(cfg.propKey);
  if (!stored) return false; // 尚未在指令碼屬性設定密碼
  return String(pwd) === String(stored);
}

/**
 * 搜尋案件（供三組通報頁面使用）。
 * keyword 可以是「班級-座號」「姓名」「案件編號」的部分文字，留空則列出所有未結案案件。
 */
function searchCases(group, pwd, keyword) {
  if (!verifyGroupPassword(group, pwd)) {
    throw new Error('密碼錯誤或登入已失效，請重新整理頁面再登入一次。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, COL.LAST_COL).getValues();
  const kw = String(keyword || '').trim();

  const matched = values.filter(function (r) {
    const classSeat = r[COL.CLASS - 1] + '-' + r[COL.SEAT - 1];
    if (!kw) return r[COL.CLOSED - 1] !== '是'; // 沒輸入關鍵字：預設列出所有未結案案件
    return (
      classSeat.indexOf(kw) !== -1 ||
      String(r[COL.NAME - 1]).indexOf(kw) !== -1 ||
      String(r[COL.CASE_ID - 1]).indexOf(kw) !== -1
    );
  });

  return matched.slice(0, 50).map(function (r) {
    return {
      caseId: r[COL.CASE_ID - 1],
      grade: r[COL.GRADE - 1],
      classNo: r[COL.CLASS - 1],
      seatNo: r[COL.SEAT - 1],
      name: r[COL.NAME - 1],
      gender: r[COL.GENDER - 1],
      symptom: r[COL.SYMPTOM - 1],
      triage: r[COL.TRIAGE - 1],
      location: r[COL.LOCATION - 1],
      hospital: r[COL.HOSPITAL - 1],
      escort: r[COL.ESCORT - 1],
      note: r[COL.NOTE - 1],
      status: r[COL.STATUS - 1],
      closed: r[COL.CLOSED - 1]
    };
  });
}

/**
 * 更新案件（供三組通報頁面使用）。
 * 依組別只允許寫入該組負責的欄位，避免誤改到其他組的資料。
 */
function updateCase(group, pwd, caseId, payload) {
  if (!verifyGroupPassword(group, pwd)) {
    throw new Error('密碼錯誤或登入已失效，請重新整理頁面再登入一次。');
  }
  const cfg = GROUP_CONFIG[group];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const row = findMainRowByCaseId_(mainSheet, caseId, COL);
  if (!row) throw new Error('找不到案件編號：' + caseId + '（可能剛好被其他人異動，請重新搜尋一次）');

  if (group === 'triage') {
    mainSheet.getRange(row, COL.TRIAGE).setValue(payload.value || '');
  } else if (group === 'rest') {
    mainSheet.getRange(row, COL.LOCATION).setValue(payload.value || '');
  } else if (group === 'hospital') {
    mainSheet.getRange(row, COL.HOSPITAL).setValue(payload.value || '');
    mainSheet.getRange(row, COL.ESCORT).setValue(payload.escort || '');
  }
  if (payload.note !== undefined) {
    mainSheet.getRange(row, COL.NOTE).setValue(payload.note);
  }
  mainSheet.getRange(row, COL.UPDATED_AT).setValue(new Date());
  mainSheet.getRange(row, COL.UPDATED_BY).setValue(cfg.label);

  syncDerivedSheets();
  return { success: true };
}

/**
 * 驗證管理者密碼（存在指令碼屬性 PWD_ADMIN）。
 */
function verifyAdminPassword(pwd) {
  const stored = PropertiesService.getScriptProperties().getProperty(ADMIN_PROP_KEY);
  if (!stored) return false;
  return String(pwd) === String(stored);
}

/**
 * 管理者：列出所有案件（含已結案），欄位齊全（含真實姓名/聯絡電話）。
 */
function adminListAllCases(pwd) {
  if (!verifyAdminPassword(pwd)) {
    throw new Error('管理者密碼錯誤或登入已失效，請重新整理頁面再登入一次。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, COL.LAST_COL).getValues();
  return values.map(function (r) {
    return {
      caseId: r[COL.CASE_ID - 1],
      grade: r[COL.GRADE - 1],
      classNo: r[COL.CLASS - 1],
      seatNo: r[COL.SEAT - 1],
      name: r[COL.NAME - 1],
      gender: r[COL.GENDER - 1],
      contact: r[COL.CONTACT - 1],
      symptom: r[COL.SYMPTOM - 1],
      triage: r[COL.TRIAGE - 1],
      location: r[COL.LOCATION - 1],
      hospital: r[COL.HOSPITAL - 1],
      escort: r[COL.ESCORT - 1],
      note: r[COL.NOTE - 1],
      status: r[COL.STATUS - 1],
      closed: r[COL.CLOSED - 1]
    };
  });
}

/**
 * 管理者：修改案件任何欄位（年級/班級/座號/姓名/性別/聯絡電話/身體症狀/
 * 檢傷初判/目前所在位置/就醫醫院/護送教師/備註）。案件編號/狀態/排序值等
 * 系統欄位不開放直接修改，會在同步時自動重新計算。
 */
function adminUpdateCase(pwd, caseId, fields) {
  if (!verifyAdminPassword(pwd)) {
    throw new Error('管理者密碼錯誤或登入已失效，請重新整理頁面再登入一次。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const row = findMainRowByCaseId_(mainSheet, caseId, COL);
  if (!row) throw new Error('找不到案件編號：' + caseId);

  const fieldColMap = {
    grade: COL.GRADE, classNo: COL.CLASS, seatNo: COL.SEAT, name: COL.NAME,
    gender: COL.GENDER, contact: COL.CONTACT, symptom: COL.SYMPTOM,
    triage: COL.TRIAGE, location: COL.LOCATION, hospital: COL.HOSPITAL,
    escort: COL.ESCORT, note: COL.NOTE
  };
  Object.keys(fieldColMap).forEach(function (key) {
    if (fields[key] !== undefined) {
      mainSheet.getRange(row, fieldColMap[key]).setValue(fields[key]);
    }
  });
  mainSheet.getRange(row, COL.UPDATED_AT).setValue(new Date());
  mainSheet.getRange(row, COL.UPDATED_BY).setValue('管理者');

  syncDerivedSheets();
  return { success: true };
}

/**
 * 管理者：新增一筆案件（不透過 Google 表單，直接手動建立）。
 */
function adminAddCase(pwd, fields) {
  if (!verifyAdminPassword(pwd)) {
    throw new Error('管理者密碼錯誤或登入已失效，請重新整理頁面再登入一次。');
  }
  if (!fields.grade || !fields.classNo || !fields.seatNo || !fields.name) {
    throw new Error('年級、班級、座號、姓名為必填欄位。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const now = new Date();
  const caseId = generateCaseId_(mainSheet, now, COL);
  const sortKey = gradeToNumber_(fields.grade) * 10000 + Number(fields.classNo) * 100 + Number(fields.seatNo);

  const newRow = new Array(COL.LAST_COL).fill('');
  newRow[COL.TIMESTAMP - 1] = now;
  newRow[COL.GRADE - 1] = fields.grade;
  newRow[COL.CLASS - 1] = fields.classNo;
  newRow[COL.SEAT - 1] = fields.seatNo;
  newRow[COL.NAME - 1] = fields.name;
  newRow[COL.GENDER - 1] = fields.gender || '';
  newRow[COL.CONTACT - 1] = fields.contact || '';
  newRow[COL.SYMPTOM - 1] = fields.symptom || '';
  newRow[COL.CASE_ID - 1] = caseId;
  newRow[COL.SORT_KEY - 1] = sortKey;
  newRow[COL.STATUS - 1] = '檢傷中';
  newRow[COL.CLOSED - 1] = '否';
  newRow[COL.UPDATED_AT - 1] = now;
  newRow[COL.UPDATED_BY - 1] = '管理者新增';

  mainSheet.appendRow(newRow);
  syncDerivedSheets();
  return { success: true, caseId: caseId };
}

/**
 * 管理者：刪除一筆案件（整列從彙整總表移除，用於演習結束後清除個資）。
 */
function adminDeleteCase(pwd, caseId) {
  if (!verifyAdminPassword(pwd)) {
    throw new Error('管理者密碼錯誤或登入已失效，請重新整理頁面再登入一次。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const row = findMainRowByCaseId_(mainSheet, caseId);
  if (!row) throw new Error('找不到案件編號：' + caseId + '（可能已經被刪除）');

  mainSheet.deleteRow(row); // 整列刪除（含被我們沒對應到的額外欄位一起刪掉，避免留下殘骸）
  syncDerivedSheets();
  return { success: true };
}

/**
 * 首頁動態表格資料（讀取去識別化總覽，不含真實姓名/電話）。
 */
function getHomeTableData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEID);
  const lastRow = sheet.getLastRow();
  const tz = Session.getScriptTimeZone();

  if (lastRow < 2) {
    return { count: 0, rows: [], updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss') };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const rows = values.map(function (r) {
    return {
      classSeat: r[0],
      name: r[1],
      gender: r[2],
      status: r[3],
      location: r[4],
      hospital: r[5],
      parentPickup: r[6],
      dischargeTime: r[7] ? Utilities.formatDate(new Date(r[7]), tz, 'yyyy/MM/dd HH:mm') : ''
    };
  });

  return {
    count: rows.length,
    rows: rows,
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss')
  };
}


// ============================================================
// 【HTML 樣板】首頁
// ============================================================
function buildHomeHtml_() {
  return '' +
'<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">' +
'<style>' +
'  * { box-sizing: border-box; }' +
'  body { margin:0; font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif; background:#f6f8f1; color:#3d4033; }' +
'  .banner { background:linear-gradient(135deg,#8fae7f,#6b9a70); padding:24px 32px; display:flex; align-items:center; gap:20px; color:#fff; }' +
'  .banner .logo { width:64px; height:64px; border-radius:50%; background:#fff; display:flex; align-items:center; justify-content:center; font-size:28px; flex-shrink:0; box-shadow:0 2px 6px rgba(0,0,0,.15); }' +
'  .banner h1 { margin:0; font-size:1.5rem; }' +
'  .banner p { margin:4px 0 0; font-size:1rem; opacity:.9; }' +
'  .container { max-width:1100px; margin:24px auto; padding:0 16px; }' +
'  .summary-bar { display:flex; align-items:center; gap:12px; justify-content:center; margin-bottom:16px; flex-wrap:wrap; }' +
'  .summary-badge { background:#4f7a5e; color:#fff; padding:10px 20px; border-radius:20px; font-weight:bold; }' +
'  .updated-at { font-size:.85rem; color:#7c8a72; }' +
'  table { width:100%; border-collapse:collapse; background:#fffdf8; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(90,120,80,.1); }' +
'  th { background:#7ba382; color:#fff; padding:10px 8px; font-size:.9rem; }' +
'  td { padding:9px 8px; text-align:center; border-bottom:1px solid #eaf1e4; font-size:.9rem; }' +
'  tr:nth-child(even) { background:#f9faf5; }' +
'  tr:last-child td { border-bottom:none; }' +
'  .status-tag { padding:3px 10px; border-radius:12px; color:#fff; font-size:.8rem; display:inline-block; }' +
'  .st-檢傷中 { background:#e0a458; } .st-休息 { background:#7fa989; } .st-送醫中 { background:#c98a5c; }' +
'  .st-已回教室 { background:#94998c; } .st-家長接回 { background:#6a8fa3; }' +
'  .empty { text-align:center; padding:40px; color:#9aa891; }' +
'  .search-bar { display:flex; align-items:center; gap:10px; justify-content:center; margin-bottom:16px; flex-wrap:wrap; }' +
'  .search-bar input[type=text], .search-bar select {' +
'    padding:8px 12px; border-radius:8px; border:1px solid #cdd8c8; font-size:.9rem; font-family:inherit;' +
'  }' +
'  .search-bar input[type=text] { width:220px; }' +
'  .search-bar button {' +
'    padding:8px 16px; border-radius:8px; border:none; background:#7ba382; color:#fff;' +
'    font-weight:bold; cursor:pointer; font-size:.9rem;' +
'  }' +
'  .search-bar button.secondary { background:#dfe6da; color:#4f5e49; }' +
'  .refresh-icon {' +
'    width:34px; height:34px; border-radius:50%; border:1px solid #cdd8c8; background:#fff;' +
'    cursor:pointer; font-size:1rem; display:flex; align-items:center; justify-content:center;' +
'  }' +
'  .refresh-icon:hover { background:#eef4ea; }' +
'</style></head><body>' +
'  <div class="banner">' +
'    <div class="logo">' + (SCHOOL_LOGO_BASE64 ? ('<img src="data:image/png;base64,' + SCHOOL_LOGO_BASE64 + '" style="width:100%;height:100%;object-fit:contain;padding:6px;">') : '<span style="font-size:28px;">&#127979;</span>') + '</div>' +
'    <div><h1>○○國小</h1><p>校園出現集體腸胃不適症狀事件通報</p></div>' + // TODO: 換成貴校校名與標語
'  </div>' +
'  <div class="container">' +
'    <div class="summary-bar">' +
'      <span class="summary-badge">學生醫療資料總覽 <span id="count">0</span> 人</span>' +
'      <span class="updated-at">最後更新：<span id="updatedAt">--</span></span>' +
'    </div>' +
'    <div class="search-bar">' +
'      <button class="refresh-icon" onclick="refresh()" title="重新整理">🔄</button>' +
'      <input type="text" id="searchInput" placeholder="搜尋姓名或班級-座號..." onkeydown="if(event.key===\\\'Enter\\\')doSearch()">' +
'      <select id="statusSelect">' +
'        <option value="">所有狀態</option>' +
'        <option value="檢傷中">檢傷中</option>' +
'        <option value="休息">休息</option>' +
'        <option value="送醫中">送醫中</option>' +
'        <option value="已回教室">已回教室</option>' +
'        <option value="家長接回">家長接回</option>' +
'      </select>' +
'      <button onclick="doSearch()">查詢</button>' +
'      <button class="secondary" onclick="showAll()">顯示全部</button>' +
'    </div>' +
'    <div id="tableWrap"><div class="empty">載入中...</div></div>' +
'  </div>' +
'<script>' +
'var homeData = { count: 0, rows: [], updatedAt: "--" };' +
'function statusClass(s){ return "status-tag st-" + s.replace(/\\(.*\\)/,""); }' +
'function escHtml(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }' +

'function passesFilter(r, kw, statusFilter){' +
'  var matchKw = !kw || String(r.classSeat).indexOf(kw) !== -1 || String(r.name||"").indexOf(kw) !== -1;' +
'  var matchStatus = !statusFilter ||' +
'    String(r.status||"").indexOf(statusFilter) === 0 ||' +
'    String(r.location||"").indexOf(statusFilter) === 0 ||' +
'    String(r.hospital||"").indexOf(statusFilter) === 0;' +
'  return matchKw && matchStatus;' +
'}' +

'function renderTable(){' +
'  var kw = document.getElementById("searchInput").value.trim();' +
'  var statusFilter = document.getElementById("statusSelect").value;' +
'  var rows = homeData.rows.filter(function(r){ return passesFilter(r, kw, statusFilter); });' +
'  var wrap = document.getElementById("tableWrap");' +
'  if (!rows.length) { wrap.innerHTML = "<div class=\\"empty\\">查無符合條件的資料</div>"; return; }' +
'  var html = "<table><thead><tr><th>班級-座號</th><th>姓名</th><th>性別</th><th>狀態</th>" +' +
'    "<th>休息觀察</th><th>送醫醫院</th><th>家長帶離</th><th>出院時間</th></tr></thead><tbody>";' +
'  rows.forEach(function(r){' +
'    html += "<tr><td>"+escHtml(r.classSeat)+"</td><td>"+escHtml(r.name)+"</td><td>"+escHtml(r.gender||"")+"</td>" +' +
'      "<td><span class=\\""+statusClass(r.status)+"\\">"+escHtml(r.status)+"</span></td>" +' +
'      "<td>"+escHtml(r.location||"")+"</td><td>"+escHtml(r.hospital||"")+"</td><td>"+escHtml(r.parentPickup||"")+"</td>" +' +
'      "<td>"+escHtml(r.dischargeTime||"")+"</td></tr>";' +
'  });' +
'  html += "</tbody></table>";' +
'  wrap.innerHTML = html;' +
'}' +

'function doSearch(){ renderTable(); }' +
'function showAll(){' +
'  document.getElementById("searchInput").value = "";' +
'  document.getElementById("statusSelect").value = "";' +
'  renderTable();' +
'}' +

'function refresh(){' +
'  google.script.run.withSuccessHandler(function(data){' +
'    homeData = data;' +
'    document.getElementById("count").textContent = data.count;' +
'    document.getElementById("updatedAt").textContent = data.updatedAt;' +
'    renderTable();' +
'  }).getHomeTableData();' +
'}' +
'refresh();' +
'setInterval(refresh, 15000);' +
'</script>' +
'</body></html>';
}


// ============================================================
// 【HTML 樣板】三組通報共用頁面（登入 + 搜尋 + 更新）
// ============================================================
function buildGroupPortalHtml_(group) {
  const cfg = GROUP_CONFIG[group];
  const optionsHtml = cfg.options.map(function (o) {
    return '<option value="' + o + '">' + o + '</option>';
  }).join('');

  return '' +
'<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">' +
'<style>' +
'  * { box-sizing: border-box; }' +
'  body { margin:0; font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif; background:' + cfg.themeColorLight + '; color:#3d3230; min-height:100vh; }' +
'  .topbar { background:' + cfg.themeColor + '; color:#3d332f; padding:16px 24px; font-size:1.2rem; font-weight:bold; }' +
'  .container { max-width:720px; margin:32px auto; padding:0 16px; }' +
'  .card { background:#fff; border-radius:12px; padding:24px; box-shadow:0 4px 16px rgba(0,0,0,.08); margin-bottom:16px; }' +
'  input, select, button, textarea { font-size:1rem; padding:10px 12px; border-radius:8px; border:1px solid #ddd; width:100%; margin-top:6px; font-family:inherit; }' +
'  label { font-weight:bold; font-size:.9rem; color:#665; }' +
'  button { background:' + cfg.themeColor + '; color:#3d332f; border:none; font-weight:bold; cursor:pointer; margin-top:12px; }' +
'  button:hover { opacity:.9; }' +
'  .msg { padding:10px; border-radius:8px; margin-top:10px; font-size:.9rem; }' +
'  .msg.error { background:#fdecea; color:#c0392b; }' +
'  .msg.ok { background:#eafaf1; color:#27ae60; }' +
'  .case-item { border:1px solid #eee; border-radius:8px; padding:12px; margin-bottom:10px; cursor:pointer; }' +
'  .case-item:hover { background:#faf9f7; }' +
'  .case-item .title { font-weight:bold; }' +
'  .case-item .status { float:right; font-size:.85rem; color:#888; }' +
'  #searchResults { max-height:400px; overflow-y:auto; }' +
'  #updatePanel { display:none; }' +
'  .hidden { display:none; }' +
'</style></head><body>' +
'  <div class="topbar">' + cfg.label + '通報系統</div>' +
'  <div class="container">' +
'    <div class="card" id="loginCard">' +
'      <label>請輸入' + cfg.label + '通報密碼</label>' +
'      <input type="password" id="pwdInput" placeholder="請輸入密碼">' +
'      <button onclick="doLogin()">登入</button>' +
'      <div id="loginMsg"></div>' +
'    </div>' +

'    <div class="card hidden" id="searchCard">' +
'      <label>搜尋學生（班級-座號 / 姓名 / 案件編號，留空列出所有未結案案件）</label>' +
'      <input type="text" id="kwInput" placeholder="例如：3-15 或 王小明">' +
'      <button onclick="doSearch()">搜尋</button>' +
'      <div id="searchMsg"></div>' +
'      <div id="searchResults"></div>' +
'    </div>' +

'    <div class="card" id="updatePanel">' +
'      <div id="updateTitle" style="font-weight:bold; margin-bottom:10px;"></div>' +
'      <label>' + cfg.updateFieldLabel + '</label>' +
'      <select id="valueSelect"><option value="">（請選擇）</option>' + optionsHtml + '</select>' +
      (group === 'hospital' ?
'      <label>護送教師</label><input type="text" id="escortInput" placeholder="護送教師姓名">' : '') +
'      <label>備註</label><textarea id="noteInput" rows="3" placeholder="選填"></textarea>' +
'      <button onclick="doUpdate()">送出更新</button>' +
'      <button style="background:#eee;" onclick="closeUpdatePanel()">取消</button>' +
'      <div id="updateMsg"></div>' +
'    </div>' +
'  </div>' +

'<script>' +
'var GROUP = "' + group + '";' +
'var PWD = "";' +
'var currentCaseId = "";' +

'function showMsg(elId, text, ok){' +
'  var el = document.getElementById(elId);' +
'  el.innerHTML = "<div class=\\"msg " + (ok ? "ok" : "error") + "\\">" + text + "</div>";' +
'}' +

'function doLogin(){' +
'  var pwd = document.getElementById("pwdInput").value;' +
'  if (!pwd) { showMsg("loginMsg", "請輸入密碼", false); return; }' +
'  google.script.run.withSuccessHandler(function(ok){' +
'    if (ok) {' +
'      PWD = pwd;' +
'      document.getElementById("loginCard").classList.add("hidden");' +
'      document.getElementById("searchCard").classList.remove("hidden");' +
'      doSearch();' +
'    } else {' +
'      showMsg("loginMsg", "密碼錯誤，請重新輸入", false);' +
'    }' +
'  }).withFailureHandler(function(err){ showMsg("loginMsg", "驗證失敗："+err.message, false); })' +
'    .verifyGroupPassword(GROUP, pwd);' +
'}' +

'function statusText(s){ return s || "檢傷中"; }' +

'var currentResults = [];' +
'function escHtml(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }' +

'function doSearch(){' +
'  var kw = document.getElementById("kwInput") ? document.getElementById("kwInput").value : "";' +
'  document.getElementById("searchResults").innerHTML = "搜尋中...";' +
'  google.script.run.withSuccessHandler(function(list){' +
'    currentResults = list;' +
'    var wrap = document.getElementById("searchResults");' +
'    if (!list.length) { wrap.innerHTML = "<div style=\\"color:#999;\\">查無符合的案件</div>"; return; }' +
'    var html = "";' +
'    list.forEach(function(c, idx){' +
'      html += "<div class=\\"case-item\\" data-idx=\\"" + idx + "\\">" +' +
'        "<span class=\\"title\\">" + escHtml(c.grade) + " " + escHtml(c.classNo) + "班" + escHtml(c.seatNo) + "號　" + escHtml(c.name) + "</span>" +' +
'        "<span class=\\"status\\">" + escHtml(statusText(c.status)) + "</span>" +' +
'        "<div style=\\"font-size:.8rem;color:#999;margin-top:4px;\\">案件編號：" + escHtml(c.caseId) + "　症狀：" + escHtml(c.symptom||"") + "</div>" +' +
'        "</div>";' +
'    });' +
'    wrap.innerHTML = html;' +
'    Array.prototype.forEach.call(wrap.querySelectorAll(".case-item"), function(el){' +
'      el.addEventListener("click", function(){ selectCase(currentResults[Number(el.getAttribute("data-idx"))]); });' +
'    });' +
'  }).withFailureHandler(function(err){ showMsg("searchMsg", "搜尋失敗："+err.message, false); })' +
'    .searchCases(GROUP, PWD, kw);' +
'}' +

'function selectCase(c){' +
'  currentCaseId = c.caseId;' +
'  document.getElementById("updateTitle").textContent =' +
'    c.grade + " " + c.classNo + "班" + c.seatNo + "號　" + c.name + "（案件編號：" + c.caseId + "）";' +
'  document.getElementById("noteInput").value = c.note || "";' +
'  if (GROUP === "triage") document.getElementById("valueSelect").value = c.triage || "";' +
'  if (GROUP === "rest") document.getElementById("valueSelect").value = c.location || "";' +
'  if (GROUP === "hospital") {' +
'    document.getElementById("valueSelect").value = c.hospital || "";' +
'    if (document.getElementById("escortInput")) document.getElementById("escortInput").value = c.escort || "";' +
'  }' +
'  document.getElementById("updatePanel").style.display = "block";' +
'  document.getElementById("updatePanel").scrollIntoView({behavior:"smooth"});' +
'}' +

'function closeUpdatePanel(){' +
'  document.getElementById("updatePanel").style.display = "none";' +
'  currentCaseId = "";' +
'}' +

'function doUpdate(){' +
'  if (!currentCaseId) { return; }' +
'  var payload = {' +
'    value: document.getElementById("valueSelect").value,' +
'    note: document.getElementById("noteInput").value' +
'  };' +
'  if (document.getElementById("escortInput")) payload.escort = document.getElementById("escortInput").value;' +
'  google.script.run.withSuccessHandler(function(){' +
'    showMsg("updateMsg", "更新成功！", true);' +
'    closeUpdatePanel();' +
'    doSearch();' +
'  }).withFailureHandler(function(err){ showMsg("updateMsg", "更新失敗："+err.message, false); })' +
'    .updateCase(GROUP, PWD, currentCaseId, payload);' +
'}' +
'</script>' +
'</body></html>';
}


// ============================================================
// 【HTML 樣板】管理者後台（新增／修改／刪除任何案件）
// ============================================================
function buildAdminHtml_() {
  const gradeOptionsHtml = GRADE_LIST.map(function (g) {
    return '<option value="' + g + '">' + g + '</option>';
  }).join('');
  const triageOptionsHtml = OPTIONS.TRIAGE.map(function (o) { return '<option value="' + o + '">' + o + '</option>'; }).join('');
  const locationOptionsHtml = OPTIONS.LOCATION.map(function (o) { return '<option value="' + o + '">' + o + '</option>'; }).join('');
  const hospitalOptionsHtml = OPTIONS.HOSPITAL.map(function (o) { return '<option value="' + o + '">' + o + '</option>'; }).join('');

  return '' +
'<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">' +
'<style>' +
'  * { box-sizing: border-box; }' +
'  body { margin:0; font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif; background:#eef2ea; color:#3d4033; min-height:100vh; }' +
'  .topbar { background:#4f7a5e; color:#fff; padding:16px 24px; font-size:1.2rem; font-weight:bold; }' +
'  .container { max-width:900px; margin:24px auto; padding:0 16px 60px; }' +
'  .card { background:#fff; border-radius:12px; padding:20px; box-shadow:0 4px 16px rgba(0,0,0,.08); margin-bottom:16px; }' +
'  input, select, button, textarea { font-size:1rem; padding:9px 10px; border-radius:8px; border:1px solid #ddd; width:100%; margin-top:6px; font-family:inherit; }' +
'  label { font-weight:bold; font-size:.85rem; color:#665; display:block; margin-top:10px; }' +
'  button { background:#4f7a5e; color:#fff; border:none; font-weight:bold; cursor:pointer; margin-top:14px; }' +
'  button.secondary { background:#ccc; color:#333; }' +
'  button.danger { background:#c0524a; color:#fff; }' +
'  .msg { padding:10px; border-radius:8px; margin-top:10px; font-size:.9rem; }' +
'  .msg.error { background:#fdecea; color:#c0392b; }' +
'  .msg.ok { background:#eafaf1; color:#27ae60; }' +
'  .row2 { display:flex; gap:10px; } .row2 > div { flex:1; }' +
'  table { width:100%; border-collapse:collapse; margin-top:10px; }' +
'  th { background:#7ba382; color:#fff; padding:8px 6px; font-size:.82rem; }' +
'  td { padding:7px 6px; text-align:center; border-bottom:1px solid #eef1ea; font-size:.82rem; }' +
'  .op-btn { display:inline-block; width:auto; margin:2px; padding:4px 10px; font-size:.8rem; }' +
'  .hidden { display:none; }' +
'  #editCard { display:none; }' +
'</style></head><body>' +
'  <div class="topbar">管理者後台</div>' +
'  <div class="container">' +

'    <div class="card" id="loginCard">' +
'      <label>請輸入管理者密碼</label>' +
'      <input type="password" id="pwdInput" placeholder="請輸入密碼">' +
'      <button onclick="doLogin()">登入</button>' +
'      <div id="loginMsg"></div>' +
'    </div>' +

'    <div class="card hidden" id="listCard">' +
'      <div style="display:flex;justify-content:space-between;align-items:center;">' +
'        <div style="font-weight:bold;">全部案件（共 <span id="totalCount">0</span> 筆）</div>' +
'        <button style="width:auto;" onclick="openAddForm()">＋ 新增案件</button>' +
'      </div>' +
'      <div id="listMsg"></div>' +
'      <div style="overflow-x:auto;"><table id="caseTable"><thead><tr>' +
'        <th>案件編號</th><th>班級-座號</th><th>姓名</th><th>性別</th><th>狀態</th><th>操作</th>' +
'      </tr></thead><tbody id="caseTableBody"></tbody></table></div>' +
'    </div>' +

'    <div class="card" id="editCard">' +
'      <div id="editTitle" style="font-weight:bold;margin-bottom:6px;"></div>' +
'      <div class="row2">' +
'        <div><label>年級</label><select id="f_grade"><option value="">（請選擇）</option>' + gradeOptionsHtml + '</select></div>' +
'        <div><label>班級</label><input type="number" id="f_classNo" min="1" max="17"></div>' +
'        <div><label>座號</label><input type="number" id="f_seatNo" min="1" max="34"></div>' +
'      </div>' +
'      <label>姓名</label><input type="text" id="f_name">' +
'      <label>性別</label><select id="f_gender"><option value="">（請選擇）</option><option value="男">男</option><option value="女">女</option></select>' +
'      <label>家人稱謂及緊急聯絡電話</label><input type="text" id="f_contact" placeholder="例如：媽媽 0911333666">' +
'      <label>身體症狀</label><input type="text" id="f_symptom" placeholder="例如：腹瀉, 嘔吐">' +
'      <label>檢傷初判</label><select id="f_triage"><option value="">（尚未檢傷）</option>' + triageOptionsHtml + '</select>' +
'      <label>目前所在位置</label><select id="f_location"><option value="">（無）</option>' + locationOptionsHtml + '</select>' +
'      <label>就醫醫院</label><select id="f_hospital"><option value="">（無）</option>' + hospitalOptionsHtml + '</select>' +
'      <label>護送教師</label><input type="text" id="f_escort">' +
'      <label>備註</label><textarea id="f_note" rows="3"></textarea>' +
'      <button onclick="submitEdit()">儲存</button>' +
'      <button class="secondary" onclick="closeEditForm()">取消</button>' +
'      <div id="editMsg"></div>' +
'    </div>' +

'  </div>' +

'<script>' +
'var PWD = "";' +
'var allCases = [];' +
'var editMode = "";' + // "edit" or "add"
'var editingCaseId = "";' +

'function showMsg(elId, text, ok){' +
'  document.getElementById(elId).innerHTML = "<div class=\\"msg " + (ok?"ok":"error") + "\\">" + text + "</div>";' +
'}' +
'function escHtml(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }' +

'function doLogin(){' +
'  var pwd = document.getElementById("pwdInput").value;' +
'  if (!pwd) { showMsg("loginMsg", "請輸入密碼", false); return; }' +
'  google.script.run.withSuccessHandler(function(ok){' +
'    if (ok) {' +
'      PWD = pwd;' +
'      document.getElementById("loginCard").classList.add("hidden");' +
'      document.getElementById("listCard").classList.remove("hidden");' +
'      loadList();' +
'    } else {' +
'      showMsg("loginMsg", "密碼錯誤，請重新輸入", false);' +
'    }' +
'  }).withFailureHandler(function(err){ showMsg("loginMsg", "驗證失敗："+err.message, false); })' +
'    .verifyAdminPassword(pwd);' +
'}' +

'function loadList(){' +
'  document.getElementById("caseTableBody").innerHTML = "<tr><td colspan=\\"6\\">載入中...</td></tr>";' +
'  google.script.run.withSuccessHandler(function(list){' +
'    allCases = list;' +
'    document.getElementById("totalCount").textContent = list.length;' +
'    var body = document.getElementById("caseTableBody");' +
'    if (!list.length) { body.innerHTML = "<tr><td colspan=\\"6\\">目前沒有任何案件資料</td></tr>"; return; }' +
'    var html = "";' +
'    list.forEach(function(c, idx){' +
'      html += "<tr><td>" + escHtml(c.caseId) + "</td>" +' +
'        "<td>" + escHtml(c.grade) + " " + escHtml(c.classNo) + "班" + escHtml(c.seatNo) + "號</td>" +' +
'        "<td>" + escHtml(c.name) + "</td><td>" + escHtml(c.gender) + "</td>" +' +
'        "<td>" + escHtml(c.status) + "</td>" +' +
'        "<td>" +' +
'          "<button class=\\"op-btn\\" data-idx=\\"" + idx + "\\" data-act=\\"edit\\">編輯</button>" +' +
'          "<button class=\\"op-btn danger\\" data-idx=\\"" + idx + "\\" data-act=\\"del\\">刪除</button>" +' +
'        "</td></tr>";' +
'    });' +
'    body.innerHTML = html;' +
'    Array.prototype.forEach.call(body.querySelectorAll("button"), function(btn){' +
'      btn.addEventListener("click", function(){' +
'        var idx = Number(btn.getAttribute("data-idx"));' +
'        var act = btn.getAttribute("data-act");' +
'        if (act === "edit") openEditForm(allCases[idx]);' +
'        if (act === "del") doDelete(allCases[idx]);' +
'      });' +
'    });' +
'  }).withFailureHandler(function(err){ showMsg("listMsg", "載入失敗："+err.message, false); })' +
'    .adminListAllCases(PWD);' +
'}' +

'function openEditForm(c){' +
'  editMode = "edit";' +
'  editingCaseId = c.caseId;' +
'  document.getElementById("editTitle").textContent = "編輯案件：" + c.caseId;' +
'  document.getElementById("f_grade").value = c.grade || "";' +
'  document.getElementById("f_classNo").value = c.classNo || "";' +
'  document.getElementById("f_seatNo").value = c.seatNo || "";' +
'  document.getElementById("f_name").value = c.name || "";' +
'  document.getElementById("f_gender").value = c.gender || "";' +
'  document.getElementById("f_contact").value = c.contact || "";' +
'  document.getElementById("f_symptom").value = c.symptom || "";' +
'  document.getElementById("f_triage").value = c.triage || "";' +
'  document.getElementById("f_location").value = c.location || "";' +
'  document.getElementById("f_hospital").value = c.hospital || "";' +
'  document.getElementById("f_escort").value = c.escort || "";' +
'  document.getElementById("f_note").value = c.note || "";' +
'  document.getElementById("editCard").style.display = "block";' +
'  document.getElementById("editCard").scrollIntoView({behavior:"smooth"});' +
'}' +

'function openAddForm(){' +
'  editMode = "add";' +
'  editingCaseId = "";' +
'  document.getElementById("editTitle").textContent = "新增案件";' +
'  ["f_grade","f_classNo","f_seatNo","f_name","f_gender","f_contact","f_symptom",' +
'   "f_triage","f_location","f_hospital","f_escort","f_note"].forEach(function(id){' +
'    document.getElementById(id).value = "";' +
'  });' +
'  document.getElementById("editCard").style.display = "block";' +
'  document.getElementById("editCard").scrollIntoView({behavior:"smooth"});' +
'}' +

'function closeEditForm(){' +
'  document.getElementById("editCard").style.display = "none";' +
'  editMode = ""; editingCaseId = "";' +
'}' +

'function collectFields(){' +
'  return {' +
'    grade: document.getElementById("f_grade").value,' +
'    classNo: document.getElementById("f_classNo").value,' +
'    seatNo: document.getElementById("f_seatNo").value,' +
'    name: document.getElementById("f_name").value,' +
'    gender: document.getElementById("f_gender").value,' +
'    contact: document.getElementById("f_contact").value,' +
'    symptom: document.getElementById("f_symptom").value,' +
'    triage: document.getElementById("f_triage").value,' +
'    location: document.getElementById("f_location").value,' +
'    hospital: document.getElementById("f_hospital").value,' +
'    escort: document.getElementById("f_escort").value,' +
'    note: document.getElementById("f_note").value' +
'  };' +
'}' +

'function submitEdit(){' +
'  var fields = collectFields();' +
'  if (editMode === "add") {' +
'    google.script.run.withSuccessHandler(function(){' +
'      showMsg("editMsg", "新增成功！", true);' +
'      closeEditForm();' +
'      loadList();' +
'    }).withFailureHandler(function(err){ showMsg("editMsg", "新增失敗："+err.message, false); })' +
'      .adminAddCase(PWD, fields);' +
'  } else {' +
'    google.script.run.withSuccessHandler(function(){' +
'      showMsg("editMsg", "更新成功！", true);' +
'      closeEditForm();' +
'      loadList();' +
'    }).withFailureHandler(function(err){ showMsg("editMsg", "更新失敗："+err.message, false); })' +
'      .adminUpdateCase(PWD, editingCaseId, fields);' +
'  }' +
'}' +

'function doDelete(c){' +
'  if (!confirm("確定要刪除「" + c.name + "（" + c.caseId + "）」這筆案件嗎？此動作無法復原！")) return;' +
'  google.script.run.withSuccessHandler(function(){' +
'    showMsg("listMsg", "已刪除", true);' +
'    loadList();' +
'  }).withFailureHandler(function(err){ showMsg("listMsg", "刪除失敗："+err.message, false); })' +
'    .adminDeleteCase(PWD, c.caseId);' +
'}' +
'</script>' +
'</body></html>';
}
