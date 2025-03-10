const settings = loadSettings();
const COLUMN_NAMES = settings.COLUMN_NAMES;

/**
 * 設定情報をGoogleスプレッドシートから読み取る関数
 * @returns {object} - 設定情報のオブジェクト
 */
function loadSettings() {
  const SETTINGS_SPREADSHEET_ID =
    "14bC3-7LaiAodEHE3OoHcbibdyLe7C5OMCM8sKQTBvLk"; // スプレッドシートIDを設定
  const sheet = SpreadsheetApp.openById(SETTINGS_SPREADSHEET_ID).getSheetByName(
    "設定"
  );
  const data = sheet.getDataRange().getValues();
  const settings = {};
  const columnNames = {};
  const columnWidths = {};

  data.forEach((row) => {
    const key = row[0];
    const value = row[1];
    try {
      if (key.startsWith("COLUMN_NAME_")) {
        const columnKey = key.replace("COLUMN_NAME_", "");
        columnNames[columnKey] = value;
      } else if (key.startsWith("COLUMN_WIDTH_")) {
        const columnKey = key.replace("COLUMN_WIDTH_", "");
        columnWidths[columnKey] = parseInt(value, 10);
      } else {
        settings[key] = JSON.parse(value);
      }
    } catch (e) {
      settings[key] = value;
    }
  });

  settings.COLUMN_NAMES = columnNames;
  settings.COLUMN_WIDTHS = columnWidths;
  return settings;
}

/**
 * メイン処理: フォルダ内のExcelファイルを結合し、Googleスプレッドシートを作成します。
 * @returns {object} - 処理結果 (成功/失敗、URL、処理ファイル数、合計行数)
 */
function combineExcelSheets() {
  try {
    const folderId = selectFolder();
    if (!folderId) throw new Error("フォルダが選択されませんでした");

    const folder = DriveApp.getFolderById(folderId);
    if (!folder) throw new Error("指定されたフォルダが見つかりません");

    const subFolders = folder.getFolders();
    const targetFolders = [];

    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      const subFolderName = subFolder.getName();
      if (
        !subFolderName.startsWith("@") &&
        subFolderName.includes("部品リスト")
      ) {
        targetFolders.push(subFolder);
      }
    }

    if (targetFolders.length === 0) {
      throw new Error("部品リストを含むフォルダが見つかりませんでした");
    }

    let combinedDataRows = [];
    let processedFiles = 0;
    const jNumberSet = new Set();
    const unitNumberSet = new Set();
    const categories = new Set();
    const filenamePartsByCategory = {
      購入: [],
      製作: [],
      電気: [],
    };

    targetFolders.forEach((targetFolder) => {
      const fileIterator = targetFolder.getFiles();
      const result = processFiles(fileIterator);
      combinedDataRows = combinedDataRows.concat(result.combinedDataRows);
      processedFiles += result.processedFiles;
      result.jNumberSet.forEach((jNumber) => jNumberSet.add(jNumber));
      result.unitNumberSet.forEach((unitNumber) =>
        unitNumberSet.add(unitNumber)
      );
      result.categories.forEach((category) => categories.add(category));
      Object.keys(result.filenamePartsByCategory).forEach((key) => {
        filenamePartsByCategory[key] = filenamePartsByCategory[key].concat(
          result.filenamePartsByCategory[key]
        );
      });
    });

    if (processedFiles === 0)
      throw new Error("処理可能なファイルが見つかりませんでした");
    if (combinedDataRows.length === 0)
      throw new Error("有効なデータが見つかりませんでした");

    const newSpreadsheet = createSpreadsheet(
      combinedDataRows,
      jNumberSet,
      unitNumberSet,
      categories,
      filenamePartsByCategory
    );
    const newSheet = newSpreadsheet.getActiveSheet();

    applyRowHeights(newSheet, combinedDataRows.length, settings.ROW_HEIGHT);
    applyConditionalFormatting(newSheet, combinedDataRows);

    return {
      success: true,
      url: newSpreadsheet.getUrl(),
      processedFiles: processedFiles,
      totalRows: combinedDataRows.length,
    };
  } catch (e) {
    console.error(`エラーが発生しました: ${e.toString()}`);
    return { success: false, error: e.toString() };
  }
}

/**
 * フォルダ内のファイルを処理する関数
 * @param {FileIterator} fileIterator - フォルダ内のファイルイテレータ
 * @returns {object} - 結合されたデータ、処理したファイル数、Jナンバーのセット、ユニットナンバーのセット
 */
function processFiles(fileIterator) {
  const validExtensions = [".xls", ".xlsx", ".xlsm"];
  let combinedDataRows = [];
  let processedFiles = 0;
  const jNumberSet = new Set();
  const unitNumberSet = new Set();
  const categories = new Set();
  const processedFileKeys = new Map();

  let filenamePartsByCategory = {
    購入: [],
    製作: [],
    電気: [],
  };

  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    const fileName = file.getName();

    if (!validExtensions.some((ext) => fileName.toLowerCase().endsWith(ext))) {
      continue; // 対応していない拡張子はスキップ
    }

    const jNumber = extractJNumber(fileName); // Jナンバー抽出関数の呼び出し
    if (jNumber) jNumberSet.add(jNumber);

    const result = processExcelFile(file, settings.TARGET_SHEET_NAME);
    if (!result) {
      console.warn(`警告: ${fileName} からデータを取得できませんでした`);
      continue;
    }
    const { fileData, category } = result;
    if (category) categories.add(category);

    // 重複チェックとファイル選択
    const { unitNumbers, categoryForCheck } = collectUnitNumbersForCheck(
      fileData,
      category
    );
    const fileKey = `${jNumber}_${categoryForCheck}_${unitNumbers}`;
    const datePrefixLength = 8; // YYYYMMDDの長さを定数化
    const currentDate = parseInt(fileName.substring(0, datePrefixLength));

    if (processedFileKeys.has(fileKey)) {
      const existingFileInfo = processedFileKeys.get(fileKey);
      const existingFile = existingFileInfo.file;
      const existingDate = existingFileInfo.date;

      if (currentDate > existingDate) {
        processedFileKeys.set(fileKey, { file: file, date: currentDate });
        combinedDataRows = combinedDataRows.filter(
          (row) => !row.includes(existingFile.getName())
        );
        combinedDataRows = combinedDataRows.concat(fileData.slice(1)); // ヘッダー行を除外
      } else {
        continue;
      }
    } else {
      processedFileKeys.set(fileKey, { file: file, date: currentDate });
      let filenamePart = processFilename(fileName, category);
      if (filenamePart !== "") {
        switch (category) {
          case "購入":
            filenamePartsByCategory["購入"].push(filenamePart);
            break;
          case "製作":
            filenamePartsByCategory["製作"].push(filenamePart);
            break;
          case "電気":
            filenamePartsByCategory["電気"].push(filenamePart);
            break;
        }
      }

      if (processedFiles === 0) {
        combinedDataRows = fileData;
      } else {
        combinedDataRows = combinedDataRows.concat(fileData.slice(1)); // ヘッダー行を除外
      }
    }
    processedFiles++;
  }

  // 各カテゴリのファイル名部分をソート
  for (let category in filenamePartsByCategory) {
    if (filenamePartsByCategory.hasOwnProperty(category)) {
      filenamePartsByCategory[category].sort((a, b) => {
        let unitA = parseInt(a.replace(/[^0-9]/g, "")) || 0;
        let unitB = parseInt(b.replace(/[^0-9]/g, "")) || 0;
        return unitA - unitB;
      });
    }
  }

  return {
    combinedDataRows,
    processedFiles,
    jNumberSet,
    unitNumberSet,
    categories,
    filenamePartsByCategory,
  };
}

/**
 * 特定のルールに基づいてファイル名を処理する関数
 * @param {string} fileName - ファイル名
 * @param {string} category - カテゴリ
 * @returns {string} - 処理されたファイル名部分
 */
function processFilename(fileName, category) {
  let filenamePart = "";

  // カテゴリに基づいてプレフィックスを設定
  let categoryPrefix = "";
  switch (category) {
    case "製作":
      categoryPrefix = "製";
      break;
    case "購入":
      categoryPrefix = "購";
      break;
    case "電気":
      categoryPrefix = "電";
      break;
    default:
      categoryPrefix = "";
  }

  // ユニット番号を抽出 (電気の場合はユニット番号なし)
  const unitNumberLength = 2; // ユニット番号の桁数
  if (category !== "電気" && categoryPrefix !== "") {
    const unitMatch = fileName.match(/(\d+)unit/i);
    if (unitMatch && unitMatch[1]) {
      const unitNumber = unitMatch[1].padStart(unitNumberLength, "0");
      filenamePart = `${categoryPrefix}${unitNumber}`;
    } else {
      filenamePart = categoryPrefix;
    }
  } else if (category === "電気") {
    filenamePart = categoryPrefix;
  }

  return filenamePart;
}

/**
 * ユニット番号を収集し、重複チェック用にフォーマットする関数
 * @param {Array<Array<string>>} data - スプレッドシートから取得したデータ
 * @param {string} category - カテゴリ
 * @returns {object} - ユニット番号とカテゴリ
 */
function collectUnitNumbersForCheck(data, category) {
  let unitNumbers = "";
  let categoryForCheck = category;
  const FIRST_DATA_ROW_INDEX = 1; // データ開始行
  for (let i = FIRST_DATA_ROW_INDEX; i < data.length; i++) {
    if (data[i][1] !== undefined && data[i][1] !== "") {
      let unitNum = String(data[i][1])
        .trim()
        .replace(/[^0-9]/g, "");
      unitNumbers = unitNum;
      break;
    }
  }
  return { unitNumbers, categoryForCheck };
}

/**
 * ファイル名からカテゴリを判定する関数
 * @param {string} fileName - ファイル名
 * @returns {string} - カテゴリ
 */
function determineCategory(fileName) {
  const match = fileName.match(/_(製作|購入|電気)_/);
  return match ? match[1] : "";
}

/**
 * Excelファイルを処理する関数
 * @param {File} file - 処理するExcelファイル
 * @param {string} targetSheetName - 対象のシート名
 * @returns {object|null} - 処理結果 (ファイルデータとカテゴリ) または null
 */
function processExcelFile(file, targetSheetName) {
  let convertedFileId = null;
  let category = "";
  const FIRST_COLUMN = 0; // A列
  const REMOVE_COLUMN_COUNT = 1; // 削除する列数
  const REQUIRED_COLUMN_COUNT = 9; // 必要な列数

  try {
    const blob = file.getBlob();
    const uploadResponse = UrlFetchApp.fetch(
      "https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart&convert=true",
      {
        method: "POST",
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        contentType: "application/vnd.ms-excel",
        payload: blob.getBytes(),
      }
    );
    convertedFileId = JSON.parse(uploadResponse.getContentText()).id;
    const convertedSheet = SpreadsheetApp.openById(convertedFileId);
    const sheet = convertedSheet.getSheetByName(targetSheetName);

    if (!sheet) {
      console.warn(
        `シート「${targetSheetName}」が見つかりません: ${file.getName()}`
      );
      return null;
    }
    const validRows = findLastDataRow(sheet);
    let data = sheet
      .getRange(1, 1, validRows, sheet.getLastColumn())
      .getValues();

    // データ整形: 不要な列の削除と必要な列数の確認
    data = data.map((row) => {
      let newRow = [...row];
      newRow.splice(FIRST_COLUMN, REMOVE_COLUMN_COUNT); // 最初の1列を削除
      return newRow.slice(0, REQUIRED_COLUMN_COUNT);
    });

    // データフィルタリング: 社内在庫なしの行を除外
    const STOCK_STATUS_COLUMN_INDEX = 2;
    const excludeStockStatus = "社内在庫なし";
    data = data.filter(
      (row, index) =>
        index === 0 || row[STOCK_STATUS_COLUMN_INDEX] !== excludeStockStatus
    );

    const fileName = file.getName();
    category = determineCategory(fileName);
    const DATE_FORMAT_COLUMN_L = 7;
    // 購入と電気の場合、L列をクリア
    if (category === "購入" || category === "電気") {
      for (let i = 1; i < data.length; i++) {
        data[i][DATE_FORMAT_COLUMN_L] = "";
      }
    }

    // カテゴリ列の追加
    data = data.map((row, index) => {
      const FIRST_ROW_INDEX = 0;
      if (index === FIRST_ROW_INDEX) {
        return [COLUMN_NAMES.D].concat(row);
      } else {
        return [category].concat(row);
      }
    });

    // 日付フォーマット
    data = data.map((row, index) => {
      const FIRST_ROW_INDEX = 0;
      const DATE_FORMAT_COLUMN_M = 8;
      const DATE_FORMAT_COLUMN_N = 9;
      if (index === FIRST_ROW_INDEX) return row;
      const dateL = row[DATE_FORMAT_COLUMN_M];
      const dateM = row[DATE_FORMAT_COLUMN_N];
      if (dateL instanceof Date) {
        row[DATE_FORMAT_COLUMN_M] = Utilities.formatDate(
          dateL,
          "JST",
          "yyyy/MM/dd"
        );
      }
      if (dateM instanceof Date) {
        row[DATE_FORMAT_COLUMN_N] = Utilities.formatDate(
          dateM,
          "JST",
          "yyyy/MM/dd"
        );
      }
      return row;
    });
    return { fileData: data, category };
  } catch (e) {
    console.error(`エラー発生 (${file.getName()}): ${e.toString()}`);
    return null;
  } finally {
    if (convertedFileId) {
      try {
        DriveApp.getFileById(convertedFileId).setTrashed(true);
      } catch (e) {
        console.error(
          `一時ファイルの削除に失敗: ${convertedFileId}, error: ${e.toString()}`
        );
      }
    }
  }
}
/**
 * シートの最終データ行を検出する関数
 * @param {Sheet} sheet - 対象のシート
 * @returns {number} - 最終データ行の行番号
 */
function findLastDataRow(sheet) {
  const lastRow = sheet.getLastRow();
  const CHECK_COLUMN = 2; // 確認する列 (B列)
  const values = sheet.getRange(1, CHECK_COLUMN, lastRow, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const cellValue = values[i][0];
    if (cellValue !== "" && cellValue !== null && cellValue !== undefined) {
      return i + 1;
    }
  }
  return 0;
}

/**
 * ファイル名からJナンバーを抽出する関数
 * @param {string} fileName - ファイル名
 * @returns {string} - 抽出されたJナンバー、存在しない場合は空文字列
 */
function extractJNumber(fileName) {
  const jNumberRegex = /(J\d{10})/; // Jナンバーの正規表現
  const match = fileName.match(jNumberRegex);
  return match ? match[1] : "";
}

/**
 * 現在の日時を文字列で取得する関数
 * @param {string} format - フォーマット文字列 (オプション)
 * @returns {string} - 現在の日時 (YYYYMMDD-HHMM形式)
 */
function getCurrentDateTime(format = "YYYYMMDD-HHMM") {
  const now = new Date();
  if (format === "YYYYMMDD-HHMM") {
    const year = now.getFullYear();
    const month = (now.getMonth() + 1).toString().padStart(2, "0");
    const day = now.getDate().toString().padStart(2, "0");
    const hours = now.getHours().toString().padStart(2, "0");
    const minutes = now.getMinutes().toString().padStart(2, "0");
    return `${year}${month}${day}-${hours}${minutes}`;
  }
}

/**
 * スプレッドシートを作成する関数
 * @param {Array<Array<string>>} combinedDataRows - 結合されたデータ（ヘッダー行を含む）
 * @param {Set<string>} jNumberSet - Jナンバーのセット
 * @param {Set<string>} unitNumberSet - ユニット番号のセット
 * @param {Set<string>} categories - カテゴリーのセット
 * @param {object} filenamePartsByCategory - カテゴリごとのファイル名部分
 * @returns {Spreadsheet} - 作成されたスプレッドシート
 */
function createSpreadsheet(
  combinedDataRows,
  jNumberSet,
  unitNumberSet,
  categories,
  filenamePartsByCategory
) {
  const DATA_START_ROW = 2; // データ開始行
  const SERIAL_NUMBER_COLUMN = 3; // 連番列 (C列)

  if (combinedDataRows.length === 0) {
    console.warn("データがないためスプレッドシートを作成しません。"); // データがない場合の警告
    return null;
  }

  const dateTime = getCurrentDateTime(); // 現在の日時を取得
  const jNumber = Array.from(jNumberSet)[0] || ""; // Jナンバー
  // ヘッダー行
  const headerRow = [
    COLUMN_NAMES.D,
    COLUMN_NAMES.E,
    COLUMN_NAMES.F,
    COLUMN_NAMES.G,
    COLUMN_NAMES.H,
    COLUMN_NAMES.I,
    COLUMN_NAMES.J,
    COLUMN_NAMES.K,
    COLUMN_NAMES.L,
    COLUMN_NAMES.M,
  ];
  let dataRows = combinedDataRows.slice(1); // データ行

  // データ行に連番、空列を追加
  let finalData = dataRows.map((row, index) =>
    ["", "", index + 1].concat(row.slice(0, 1).concat(row.slice(1)))
  );

  // ヘッダー行を追加
  finalData.unshift(
    [COLUMN_NAMES.A, COLUMN_NAMES.B, COLUMN_NAMES.C].concat(headerRow)
  );

  const unitNumbersStr = Array.from(unitNumberSet)
    .filter((unit) => unit !== "")
    .sort()
    .join("-"); // ユニット番号

  // カテゴリごとにファイル名部分を結合
  let combinedFilenameBase = "";
  let filenameParts = [];
  if (filenamePartsByCategory["購入"].length > 0) {
    filenameParts.push(
      "購" +
        filenamePartsByCategory["購入"].map((part) => part.slice(1)).join("-")
    );
  }
  if (filenamePartsByCategory["製作"].length > 0) {
    filenameParts.push(
      "製" +
        filenamePartsByCategory["製作"].map((part) => part.slice(1)).join("-")
    );
  }
  if (filenamePartsByCategory["電気"].length > 0) {
    filenameParts.push("電");
  }
  combinedFilenameBase = filenameParts.join("_");

  // combinedFilenameBaseが空でない場合にアンダースコアを追加
  let underscore = combinedFilenameBase !== "" ? "_" : "";

  // ファイル名に元のファイル名の一部を追加
  const newFileName = `${jNumber}${underscore}${combinedFilenameBase}_${dateTime}`;

  // テンプレートスプレッドシートをコピー
  const newSpreadsheet = copySpreadsheet(
    settings.TEMPLATE_SPREADSHEET_ID,
    newFileName,
    settings.OUTPUT_FOLDER_ID
  );

  const newSheet = newSpreadsheet.getActiveSheet();

  const range = newSheet.getRange(1, 1, finalData.length, finalData[0].length);
  range.setValues(finalData);

  newSheet.setFrozenRows(1); // ヘッダー行を固定

  // ヘッダー行の書式設定
  const headerRange = newSheet.getRange(1, 1, 1, finalData[0].length);
  headerRange
    .setBackground("#f3f3f3")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // 連番ヘッダーの書式設定
  const serialNumberRangeHeader = newSheet.getRange(
    1,
    SERIAL_NUMBER_COLUMN,
    finalData.length,
    1
  );
  serialNumberRangeHeader.setHorizontalAlignment("right");

  // Unit番号列の書式設定
  const unitRange = newSheet.getRange(1, 5, finalData.length, 1); // E列
  unitRange.setHorizontalAlignment("right");

  // 部品番号列の書式設定
  const partNumberRange = newSheet.getRange(
    2,
    6,
    combinedDataRows.length - 1,
    1
  ); // F列
  partNumberRange.setNumberFormat("000");

  // チェックボックスの設定
  const checkboxRange1 = newSheet.getRange(2, 1, finalData.length - 1, 1); // A列
  const checkboxRange2 = newSheet.getRange(2, 2, finalData.length - 1, 1); // B列
  checkboxRange1.insertCheckboxes();
  checkboxRange2.insertCheckboxes();

  // 列幅の設定
  Object.entries(settings.COLUMN_WIDTHS).forEach(([col, width]) => {
    const columnIndex = col.charCodeAt(0) - "A".charCodeAt(0) + 1;
    newSheet.setColumnWidth(columnIndex, width);
  });

  // カテゴリ列の書式設定
  const categoryRange = newSheet.getRange(1, 4, finalData.length, 1); // D列
  categoryRange.setHorizontalAlignment("center");

  // Unit番号列の書式設定(数値フォーマット)
  const unitNumberRange = newSheet.getRange(
    2,
    5,
    combinedDataRows.length - 1,
    1
  ); // E列
  unitNumberRange.setNumberFormat('00"unit"');

  // 部品番号列の数値フォーマット
  const partNumberRangeWithHyphen = newSheet.getRange(
    2,
    6,
    combinedDataRows.length - 1,
    1
  );
  partNumberRangeWithHyphen.setNumberFormat('"-"000');

  // 連番の計算式を設定
  const serialNumberFormulaRange = newSheet.getRange(
    2,
    SERIAL_NUMBER_COLUMN,
    finalData.length - 1,
    1
  );
  serialNumberFormulaRange.setFormula("=ROW()-1");
  serialNumberFormulaRange.setNumberFormat("@"); // 書式を文字列に設定

  // ソート
  const sortRange = newSheet.getRange(
    2,
    1,
    finalData.length - 1,
    finalData[0].length
  );
  sortRange.sort([
    { column: 4, ascending: true }, // カテゴリ
    { column: 5, ascending: true }, // Unit番号
    { column: 11, ascending: true }, // 手配先
    { column: 6, ascending: true }, // 部品番号
  ]);

  applyAlternatingRowColors(newSheet, finalData.length);

  // カテゴリ別背景色
  for (let i = DATA_START_ROW; i <= finalData.length; i++) {
    const category = newSheet.getRange(i, 4).getValue();
    const range2 = newSheet.getRange(i, 9); // I列
    if (category === "製作") range2.setBackground("#e0ffff");
    else if (category === "購入") range2.setBackground("#ffe4e1");
    else if (category === "電気") range2.setBackground("#ffff00");
  }
  applyUnitColors(newSheet, finalData);

  // 連番の計算式を値に変換
  const lastRow = newSheet.getLastRow();

  const serialNumberRange = newSheet.getRange(
    2,
    SERIAL_NUMBER_COLUMN,
    lastRow - 1,
    1
  );
  serialNumberRange.copyValuesToRange(
    newSheet,
    SERIAL_NUMBER_COLUMN,
    SERIAL_NUMBER_COLUMN,
    2,
    lastRow
  );

  return newSpreadsheet;
}

/**
 * テンプレートスプレッドシートをコピーする関数
 * @param {string} templateId - テンプレートスプレッドシートのID
 * @param {string} newFileName - 新しいファイル名
 * @param {string} outputFolderId - 出力フォルダのID
 * @returns {Spreadsheet} - コピーされたスプレッドシート
 */
function copySpreadsheet(templateId, newFileName, outputFolderId) {
  const file = DriveApp.getFileById(templateId).makeCopy(newFileName);
  const outputFolder = DriveApp.getFolderById(outputFolderId);
  outputFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return SpreadsheetApp.openById(file.getId());
}

/**
 * 行の高さを適用し、文字位置を中央に設定する関数
 * @param {Sheet} sheet - 対象のシート
 * @param {number} rowCount - 行数
 * @param {number} rowHeight - 行の高さ
 */
function applyRowHeights(sheet, rowCount, rowHeight) {
  sheet.setRowHeights(2, rowCount - 1, rowHeight);
  const range = sheet.getRange(2, 1, rowCount - 1, sheet.getLastColumn());
  range.setVerticalAlignment("middle");
}
/**
 * 行に交互の背景色を適用する関数
 * @param {Sheet} sheet - 対象のシート
 * @param {number} rowCount - 行数
 */
function applyAlternatingRowColors(sheet, rowCount) {
  if (rowCount <= 1) return;
  const lastColumn = sheet.getLastColumn();
  const range = sheet.getRange(2, 1, rowCount - 1, lastColumn);
  const backgrounds = [];

  for (let i = 0; i < rowCount - 1; i++) {
    const rowColor = i % 2 === 0 ? "#ffffff" : "#f2f2f2";
    backgrounds.push(Array(lastColumn).fill(rowColor));
  }

  range.setBackgrounds(backgrounds);
}

/**
 * 条件付き書式を適用する関数
 * @param {Sheet} sheet - 対象のシート
 * @param {Array<Array<string>>} data - データ配列
 */
function applyConditionalFormatting(sheet, data) {
  if (!data || data.length <= 1) {
    return;
  }

  const categoryCol = 4;
  const supplierCol = 11;
  const processingCol = 10;
  const dateCols = [12, 13];
  const lastColumn = sheet.getLastColumn();

  try {
    const rules = sheet.getConditionalFormatRules();
    rules.forEach((rule) => sheet.removeConditionalFormatRule(rule));
  } catch (e) {
    console.warn("条件付き書式のクリアに失敗しました。");
  }

  const checkedBackgroundColorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$A2=TRUE`)
    .setBackground("#a9a9a9")
    .setRanges([sheet.getRange(2, 1, data.length - 1, lastColumn)])
    .build();

  const checkedTextColorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$B2=TRUE`)
    .setFontColor("#ff0000")
    .setRanges([sheet.getRange(2, 1, data.length - 1, lastColumn)])
    .build();

  const categoryRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("製作")
      .setBackground("#add8e6")
      .setRanges([sheet.getRange(2, categoryCol, data.length - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("購入")
      .setBackground("#ffc0cb")
      .setRanges([sheet.getRange(2, categoryCol, data.length - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("電気")
      .setBackground("#ffff00")
      .setRanges([sheet.getRange(2, categoryCol, data.length - 1, 1)])
      .build(),
  ];

  const supplierRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("社内")
      .setBackground("#90ee90")
      .setRanges([sheet.getRange(2, supplierCol, data.length - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("MISUMI")
      .setBackground("#ffff00")
      .setRanges([sheet.getRange(2, supplierCol, data.length - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("KEYENCE")
      .setBackground("#d3d3d3")
      .setRanges([sheet.getRange(2, supplierCol, data.length - 1, 1)])
      .build(),
  ];

  const processingRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("塗装")
      .setBackground("#add8e6")
      .setRanges([sheet.getRange(2, processingCol, data.length - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("ユニクロ")
      .setBackground("#ffc0cb")
      .setRanges([sheet.getRange(2, processingCol, data.length - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("アルマイト")
      .setBackground("#90ee90")
      .setRanges([sheet.getRange(2, processingCol, data.length - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("硬質クロム")
      .setBackground("#dda0dd")
      .setRanges([sheet.getRange(2, processingCol, data.length - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("無電解ニッケル")
      .setBackground("#ffa500")
      .setRanges([sheet.getRange(2, processingCol, data.length - 1, 1)])
      .build(),
  ];

  const dateRules = [];
  dateCols.forEach((col) => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=INDIRECT("RC",FALSE)<TODAY()`)
      .setFontColor("#ff0000")
      .setRanges([sheet.getRange(2, col, data.length - 1, 1)])
      .build();
    dateRules.push(rule);
    const dateRange = sheet.getRange(2, col, data.length - 1, 1);
    dateRange.setNumberFormat("mm/dd_aaa");
  });

  sheet.setConditionalFormatRules(
    [checkedBackgroundColorRule, checkedTextColorRule]
      .concat(categoryRules)
      .concat(supplierRules)
      .concat(processingRules)
      .concat(dateRules)
  );
}

/**
 * ユニット単位で背景色を適用する関数
 * @param {Sheet} sheet - 対象のシート
 * @param {Array<Array<string>>} data - データ配列
 */
function applyUnitColors(sheet, data) {
  const UNIT_COLUMN = 5; // E列
  const DATA_START_ROW = 2;

  if (data.length <= 1) return;
  let lastUnit = null;
  let colorFlag = false;
  const ranges = [];
  const backgrounds = [];

  for (let i = DATA_START_ROW; i <= data.length; i++) {
    const currentUnit = data[i - 1][UNIT_COLUMN - 1];
    if (currentUnit !== lastUnit) {
      colorFlag = !colorFlag;
      lastUnit = currentUnit;
    }
    ranges.push(i);
    backgrounds.push(colorFlag ? "#d3d3d3" : "#f0f0f0");
  }

  const range = sheet.getRange(DATA_START_ROW, UNIT_COLUMN, data.length - 1, 1);
  const backgroundColors = range.getBackgrounds();

  for (let i = 0; i < ranges.length; i++) {
    const rowIndex = ranges[i] - DATA_START_ROW;
    backgroundColors[rowIndex][0] = backgrounds[i];
  }

  range.setBackgrounds(backgroundColors);
}
/**
 * 複数のセルに背景色を設定する関数
 * @param {Sheet} sheet - 対象のシート
 * @param {Array<Array<number|string>>} ranges - [行番号, 列番号, 色]の配列
 */
function setBackgroundColors(sheet, ranges) {
  const backgrounds = ranges.map(([row, col, color]) => {
    return [color];
  });

  const range = sheet.getRange(ranges[0][0], ranges[0][1], ranges.length, 1);
  range.setBackgrounds(backgrounds);
}

/**
 * ユーザーに指定されたフォルダ内のサブフォルダ選択ダイアログを表示し、選択されたフォルダのIDを取得する関数
 * @returns {string} - 選択されたフォルダのID
 */
function selectFolder() {
  const parentFolderId = "1v9HrinPjhW-ey5WSrX8GNouyt3eWDqxR"; // 指定されたフォルダのID
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const folders = parentFolder.getFolders();
  const folderNames = [];
  const folderIds = [];

  while (folders.hasNext()) {
    const folder = folders.next();
    const folderName = folder.getName();
    if (!folderName.startsWith("@")) {
      folderNames.push(folderName);
      folderIds.push(folder.getId());
    }
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "フォルダ選択",
    "以下のフォルダから選択してください:\n" + folderNames.join("\n"),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const folderName = response.getResponseText();
    const folderIndex = folderNames.indexOf(folderName);
    if (folderIndex !== -1) {
      return folderIds[folderIndex];
    } else {
      ui.alert("無効なフォルダ名です。");
      return null;
    }
  } else {
    ui.alert("フォルダ選択がキャンセルされました。");
    return null;
  }
}

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("変換メニュー")
    .addItem("対象フォルダの選択", "combineExcelSheets")
    .addToUi();
}
