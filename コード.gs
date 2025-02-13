// 定数定義
const SHEET_NAME_LIST = '一覧表';
const SHEET_NAME_SETTINGS = '設定値';

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
    const settings = getSettingData(spreadsheet);
    const { invoiceFolderName, promptNotes, batchSize, sleepDuration, apiEndpoint, unprocessedFolderName, processedFolderName } = settings;
    
    // 処理ステータス表示
    displayStatus(spreadsheet, '処理対象PDF数を取得しています。');
    
    // 処理対象取得
    const pdfFiles = getPDFFilesInFolder(`${invoiceFolderName}/${unprocessedFolderName}`);
    displayStatus(spreadsheet, `処理対象PDF数を取得しました。処理対象の件数は ${pdfFiles.length} 件です。`);
    
    // API呼び出し＆スプレッドシート書き込み＆フォルダ移動
    const { successCount, errorCount } = processPDFFiles(
      spreadsheet, pdfFiles, promptNotes,
      `${invoiceFolderName}/${processedFolderName}`,
      batchSize, sleepDuration, apiEndpoint
    );
    displayStatus(spreadsheet, `処理完了しました。 ( 正常: ${successCount} 件 + エラー: ${errorCount} 件) / 全: ${pdfFiles.length} 件`);
  } catch (error) {
    Logger.log(`メイン処理中にエラーが発生しました: ${error}`);
    displayStatus(spreadsheet, "main() 処理中にエラーが発生しました。", error);
  }
}

// ステータス表示
function displayStatus(spreadsheet, message, error = null) {
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
  const settings = {
    invoiceFolderName:     settingsSheet.getRange('C2').getValue(), // フォルダ名
    promptNotes:           settingsSheet.getRange('C3').getValue(), // プロンプトのメモ
    batchSize:             settingsSheet.getRange('C4').getValue(), // バッチサイズ
    sleepDuration:         settingsSheet.getRange('C5').getValue(), // スリープ時間
    apiEndpoint:           settingsSheet.getRange('C6').getValue(), // APIエンドポイント
    unprocessedFolderName: settingsSheet.getRange('C7').getValue(), // 「未処理」フォルダ名
    processedFolderName:   settingsSheet.getRange('C8').getValue()  // 「処理済」フォルダ名
  };
  return settings;
}

// フォルダ内のPDFファイル取得
function getPDFFilesInFolder(folderName) {
  const folder = getFolderByPath(folderName);

  if (!folder) {
    Logger.log(`フォルダ "${folderName}" が見つかりません。`);
    throw new Error(`フォルダ "${folderName}" が見つかりません。`);
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

// PDFファイル処理（バッチサイズごとに分割）
function processPDFFiles(spreadsheet, pdfFiles, promptNotes, destFolderName, batchSize, sleepDuration, apiEndpoint) {
  let successCount = 0; // 正常件数
  let errorCount = 0;   // エラー件数
  const dataToWrite = []; // 書き込むデータを蓄積する二次元配列
  const filesToMove = []; // 移動するファイルを蓄積する配列

  try {
    // スプレッドシートの書き込み済の最終行を取得
    const sheet = spreadsheet.getSheetByName(SHEET_NAME_LIST);
    let lastRow = sheet.getRange(sheet.getMaxRows(), 6).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

    // 処理済フォルダの取得
    const destFolder = getFolderByPath(destFolderName);
    if (!destFolder) {
      Logger.log(`フォルダ "${destFolderName}" が見つかりません。`);
      throw new Error(`フォルダ "${destFolderName}" が見つかりません。`);
    }

    // バッチサイズごとに処理するためのループ
    for (let i = 0; i < pdfFiles.length; i += batchSize) {
      const batchFiles = pdfFiles.slice(i, i + batchSize); // バッチサイズごとに取得

      // バッチごとの処理
      batchFiles.forEach((pdfFile, index) => {
        const globalIndex = i + index + 1; // 全体のインデックス
        displayStatus(spreadsheet, `"${pdfFile.getName()}" を処理中です。（ ${globalIndex} / ${pdfFiles.length} 件目）`);
        logFileDetails(pdfFile, globalIndex);
        Utilities.sleep(sleepDuration); // スリープを入れることでAPIのレート制限を回避

        const invoiceData = extractInvoiceDataFromPDF(pdfFile.getId(), promptNotes, apiEndpoint, spreadsheet);
        if (invoiceData) {
          logInvoiceData(invoiceData);
          // データを二次元配列に追加
          dataToWrite.push([
            invoiceData.date, invoiceData.issuer, invoiceData.amount,
            "", "", // 空のセル
            pdfFile.getUrl(), pdfFile.getName()
          ]);
          filesToMove.push(pdfFile); // 移動するファイルを配列に追加
          successCount++; // 正常件数をカウント
        } else {
          Logger.log('請求書データを取得できませんでした。');
          errorCount++; // エラー件数をカウント
        }
      });

      // バッチごとにスプレッドシートに書き込み
      if (dataToWrite.length > 0) {
        writeDataToSheet(spreadsheet, lastRow + 1, dataToWrite);
        lastRow += dataToWrite.length;
        dataToWrite.length = 0; // データをクリア
      }

      // バッチごとにファイルを移動
      filesToMove.forEach((pdfFile) => {
        pdfFile.moveTo(destFolder);
      });
      filesToMove.length = 0; // 移動するファイルをクリア
    }
  } catch (error) {
    Logger.log(`PDFファイル処理中にエラーが発生しました: ${error}`);
    displayStatus(spreadsheet, "processPDFFiles() 処理中にエラーが発生しました。", error);
  }

  return { successCount, errorCount }; // 正常件数とエラー件数を返す
}

// PDFから請求書データを抽出
function extractInvoiceDataFromPDF(fileId, notes, apiEndpoint, spreadsheet) {
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

    const response = UrlFetchApp.fetch(apiEndpoint, options);
    const data = JSON.parse(response.getContentText());

    return data.invoice_data;
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error}`);
    displayStatus(spreadsheet, "extractInvoiceDataFromPDF() 処理中にエラーが発生しました。", error);
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
