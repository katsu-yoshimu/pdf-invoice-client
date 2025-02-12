// 定数定義
const SHEET_NAME_LIST = '一覧表';
const SHEET_NAME_SETTINGS = '設定値';
const API_ENDPOINT = 'https://vercel-sandbox-git-main-katsu-yoshimus-projects.vercel.app/api/exctract_invoice/';
const SLEEP_DURATION = 4000; // APIのレート制限を回避するためのスリープ時間（ミリ秒）

// 実行ダイアログ、結果ダイアログ
function showExecutionDialogs() {
  var ui = SpreadsheetApp.getUi(); // UIを取得

  // 実行確認ダイアログ
  var response = ui.alert(
    '確認',
    '請求書PDFの読込を実行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    // 実行処理（ここに実際の処理を記述）
    main();

    // 実行完了ダイアログ
    ui.alert('請求書PDFの読込が完了しました');
  } else {
    ui.alert('請求書PDFの読込がキャンセルされました');
  }
}

// メイン処理
function main() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // スプレッドシートの参照を一度取得
  try {  
    // 設定値取得
    const { invoiceFolderName, promptNotes } = getSettingData(spreadsheet);
    // 処理ステータス表示
    displayStatus(spreadsheet, '処理対象PDF数を取得しています。');

    // 処理対象取得
    const pdfFiles = getPDFFilesInFolder(invoiceFolderName + '/未処理');
    displayStatus(spreadsheet, `処理対象PDF数を取得しました。処理対象の件数は ${pdfFiles.length} 件です。`);

    // API呼び出し＆スプレッドシート書き込み＆フォルダ移動
    const { successCount, errorCount } = processPDFFiles(spreadsheet, pdfFiles, promptNotes, invoiceFolderName + '/処理済');
    displayStatus(spreadsheet, `処理完了しました。正常件数: ${successCount} 件, エラー件数: ${errorCount} 件`);
  } catch (error) {
    Logger.log(`メイン処理中にエラーが発生しました: ${error}`);
    displayStatus(spreadsheet, "main() 処理中にエラーが発生しました。", error);
  }
}

// ステータス表示
function displayStatus(spreadsheet, message, error = null) {
  // 今日の日付を表示
  var date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss - ');

  const sheet = spreadsheet.getSheetByName(SHEET_NAME_LIST);
  if (!error) {
    sheet.getRange(1, 3).setValue(date + message);
  } else {
    sheet.getRange(2, 3).setValue(date + message + `”${error}"`);
  }
  SpreadsheetApp.flush();
}

// 設定データ取得
function getSettingData(spreadsheet) {
  const settingsSheet = spreadsheet.getSheetByName(SHEET_NAME_SETTINGS);
  const invoiceFolderName = settingsSheet.getRange('C2').getValue();
  const promptNotes = settingsSheet.getRange('C3').getValue();
  return { invoiceFolderName, promptNotes };
}

// フォルダ内のPDFファイル取得
function getPDFFilesInFolder(folderName) {
  const folder = getFolderByPath(folderName);

  if (!folder) {
    Logger.log(`フォルダ "${folderName}" が見つかりません。`);
    throw new Error(`フォルダ "${folderName}" が見つかりません。`);
    // return [];
  }

  return getPDFFilesSortedByName(folder);
}

// フォルダ内のPDFファイルを名前順にソートして取得
function getPDFFilesSortedByName(folder) {
  const files = folder.getFiles();
  const pdfFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === 'application/pdf') {
      pdfFiles.push(file);
    }
  }

  pdfFiles.sort((a, b) => a.getName().localeCompare(b.getName()));
  return pdfFiles;
}

// PDFファイル処理
function processPDFFiles(spreadsheet, pdfFiles, promptNotes, destFolderName) {
  let successCount = 0; // 正常件数
  let errorCount = 0;   // エラー件数
  const dataToWrite = []; // 書き込むデータを蓄積する二次元配列
  const filesToMove = []; // 移動するファイルを蓄積する配列

  try {
    // スプレッドシートの書き込み済の最終行を取得
    const sheet = spreadsheet.getSheetByName(SHEET_NAME_LIST);
    const lastRow = sheet.getRange(sheet.getMaxRows(), 6).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    // 処理済フォルダの取得
    const destFolder = getFolderByPath(destFolderName);
    if (!destFolder) {
      Logger.log(`フォルダ "${destFolderName}" が見つかりません。`);
      throw new Error(`フォルダ "${destFolderName}" が見つかりません。`);
      // return { successCount, errorCount };
    }

    pdfFiles.forEach((pdfFile, index) => {
      displayStatus(spreadsheet, `"${pdfFile.getName()}" を処理中です。（ ${index + 1} / ${pdfFiles.length} 件目）`);
      logFileDetails(pdfFile, index + 1);
      Utilities.sleep(SLEEP_DURATION); // スリープを入れることでAPIのレート制限を回避

      const invoiceData = extractInvoiceDataFromPDF(pdfFile.getId(), promptNotes);
      if (invoiceData) {
        logInvoiceData(invoiceData);
        // データを二次元配列に追加
        dataToWrite.push([
          invoiceData.date,
          invoiceData.issuer,
          invoiceData.amount,
          "", // 空のセル
          "", // 空のセル
          pdfFile.getUrl(),
          pdfFile.getName()
        ]);
        filesToMove.push(pdfFile); // 移動するファイルを配列に追加
        successCount++; // 正常件数をカウント
      } else {
        Logger.log('請求書データを取得できませんでした。');
        errorCount++; // エラー件数をカウント
      }
    });

    // すべてのデータを一括で書き込む
    if (dataToWrite.length > 0) {
      writeDataToSheet(spreadsheet, lastRow + 1, dataToWrite);

      // データの書き込みが成功した後にファイルを移動
      filesToMove.forEach((pdfFile) => {
        pdfFile.moveTo(destFolder);
      });
    }
  } catch (error) {
    Logger.log(`PDFファイル処理中にエラーが発生しました: ${error}`);
    displayStatus(spreadsheet, "processPDFFiles() 処理中にエラーが発生しました。", error);
  }

  return { successCount, errorCount }; // 正常件数とエラー件数を返す
}

// PDFから請求書データを抽出
function extractInvoiceDataFromPDF(fileId, notes) {
  try {
    const file = DriveApp.getFileById(fileId);
    const fileBlob = file.getBlob();

    const payload = {
      'file': fileBlob,
      'notes': notes,
    };

    const options = {
      'method': 'post',
      'contentType:': 'multipart/form-data',
      'payload': payload,
      // 'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    const data = JSON.parse(response.getContentText());

    return data.invoice_data;
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error}`);
    return null;
  }
}

// スプレッドシートにデータを一括書き込み
function writeDataToSheet(spreadsheet, startRow, data) {
  const sheet = spreadsheet.getSheetByName(SHEET_NAME_LIST);

  // データを一括で書き込む
  sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.flush();
}

// ファイル詳細ログ
function logFileDetails(pdfFile, fileNumber) {
  Logger.log(`<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< No: ${fileNumber} >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>`);
  Logger.log(`File Name : ${pdfFile.getName()}`);
  Logger.log(`URL       : ${pdfFile.getUrl()}`);
  Logger.log(`FILE ID   : ${pdfFile.getId()}`);
}

// 請求書データログ
function logInvoiceData(invoiceData) {
  Logger.log(`日付      : ${invoiceData.date}`);
  Logger.log(`請求元    : ${invoiceData.issuer}`);
  Logger.log(`金額      : ${invoiceData.amount}`);
}

// フォルダ有無チェック
function getFolderByPath(folderPath) {
  var pathParts = folderPath.split('/');

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();

  for (var i = 0; i < pathParts.length; i++) {
    var folderName = pathParts[i];
    var folders = currentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
      currentFolder = folders.next(); // 次のフォルダへ移動
    } else {
      Logger.log(`フォルダ "${folderName}" が見つかりません。`);
      return null; // 指定のフォルダが見つからなかった場合
    }
  }
  return currentFolder;
}

// テスト関数群
function test_getFolderByPath() {
  var folderPath = '25年度インボイス/未処理'; // フォルダパス
  var folder = getFolderByPath(folderPath);

  if (folder) {
    Logger.log('フォルダ "' + folder.getName() + '" を取得しました。');
  } else {
    Logger.log('フォルダ "' + folderPath + '" は存在しません。');
  }
}

function test_api() {
  try {
    // API呼び出し
    var fileId = '1ly1ACBdKrWFJG2VLCCkr45WIHt4r3L1l'
    var prompt_notes = ''
    var invoice_data = extractInvoiceDataFromPDF(fileId, prompt_notes)
    // 7. 請求書データを表示
    if (invoice_data) {
      Logger.log('日付: ' + invoice_data.date);
      Logger.log('請求元: ' + invoice_data.issuer);
      Logger.log('金額: ' + invoice_data.amount);
    } else {
      Logger.log('請求書データを取得できませんでした。');
      //Logger.log('エラー内容: ' + data.gemini_response.text); // エラー内容も確認
    }
  } catch (error) {
    Logger.log(`APIテスト中にエラーが発生しました: ${error}`);
  }
}