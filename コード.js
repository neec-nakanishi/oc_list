//
// 上の【 ▷実行 】ボタンを押すと処理が開始されます
//

const YEAR = "2026";

const departments = [
  "クリエイターズ",
  "デザイン",
  "ミュージック",
  "IT",
  "テクノロジー",
  "スポーツ・医療"
];

// コピーしたい列を指定（A列とC列 → 列番号1と3　※ 1始まり）
const columnsArray = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 21, 22, 23, 29, 30, 32, 33, 40, 41];

function run() {
  // メッセージをクリア
  resetMessage();
  
  // 出力ファイル名の設定
  var createFileName = getTodayAsMMDD();
  // 蒲田用
  var srcFolderId = '1SxwW5EZ2ZY36DLKrLT18DIaaaWEvWomk';  // ← 参加者リストアップロードフォルダのIDを指定
  var trgFolderId = '1olld1vrLi0WyamgWkWYpIx-8OUM3SWGR';  // ← 出力先フォルダのIDを指定
  // 実績出力のチェック
  const fixF = SpreadsheetApp.getActiveSheet().getRange(5, 5).getValue();
  if (fixF) {
    // 出力先を「蒲田」ー「実施後」フォルダに変更
    trgFolderId = '1LhAOs8z9iEkTrPPDqlG9yeIcvgCv-aAP';
    // 出力ファイル名に"_fix"を追加
    createFileName = '_fix';
  }
  // ファイル作成実行
  var ret = ReadSpreadsheet(srcFolderId, trgFolderId, createFileName);
  if (ret) {
    // 完了メッセージ
    message("蒲田校リストの作成が完了しました", 11);
  }

  // 八王子用
  srcFolderId = '1NexF-3J2auuge4EjwKPFAWdiPEt5-Ray';  // ← 参加者リストアップロードフォルダのIDを指定
  trgFolderId = '1SEHKMJkKwXYtfBu_6eC2YbmdVGLo8b5d';  // ← 出力先フォルダのIDを指定
  // 実績出力のチェック
  if (fixF) {
    // 出力先を「八王子」ー「実施後」フォルダに変更
    trgFolderId = '1NtTDy2h2U-J7m-tvwfK2anXAQXNGPfRB';
  }

  ret = ReadSpreadsheet(srcFolderId, trgFolderId, createFileName);
  if (ret) {
    // 完了メッセージ
    message("八王子校リストの作成が完了しました", 12);
  }

  // 実績出力のチェックをリセットする
  SpreadsheetApp.getActiveSheet().getRange(5,5).setValue(false);

}

// 参加者リストの読み込み
function ReadSpreadsheet(srcFolderId, trgFolderId, createFileName) {
  var ret = false;
  const srcFolder = DriveApp.getFolderById(srcFolderId);
  var files = srcFolder.getFiles();
  var rmfiles = [];

  while (files.hasNext()) {
    const file = files.next();
    // xlsxファイル以外はパスする
    if (file.getMimeType() != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
      continue;
    }
    // Excelからスプレッドシートに一時的に変換
    const tmpFile = Drive.Files.create({
      name: createFileName,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [srcFolderId]
    }, file.getBlob(), {convert: true, supportsAllDrives: true});
    Logger.log("読み込んだファイル: " + file.getName());
    // 削除ファイルリストに追加
    rmfiles.push(file.getId());
    rmfiles.push(tmpFile.getId());

    try {
      const sourceSheet = SpreadsheetApp.openById(tmpFile.id);
      copySelectedColumnsToNewSpreadsheet(srcFolder, sourceSheet, trgFolderId, createFileName);
      ret = true;
    } catch(e) {
      errorMessage('エラー内容：'+e.message);
      console.log('エラー内容：'+e.message);
    }
  }

  try {
    // リストアップされたファイルをすべて削除
    files = srcFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (rmfiles.includes(file.getId())) {
        file.setTrashed(true);
      }
    }
  } catch(e) {
    errorMessage('ファイル削除　エラー内容：'+e.message);
    console.log('ファイル削除　エラー内容：'+e.message);
  }
  return ret;
}

function resetMessage() {
  message("", 11);
  message("", 12);
  errorMessage("");
}

function message(str, row) {
  SpreadsheetApp.getActiveSheet().getRange(row,1).setFontColor("#0000ff");
  SpreadsheetApp.getActiveSheet().getRange(row,1).setValue(str);
}

function errorMessage(str) {
  SpreadsheetApp.getActiveSheet().getRange(10,1).setFontColor("#ff0000");
  SpreadsheetApp.getActiveSheet().getRange(10,1).setValue(str);
}


// 列を指定して学科配布用参加者リストを作成
function copySelectedColumnsToNewSpreadsheet(srcFolder, sourceSheet, trgFolderId, createFileName) {
  // コピーデータの作成
  const data = sourceSheet.getDataRange().getValues();
  const copiedData = [];
  for (let i = 0; i < data.length; i++) {
    const row = columnsArray.map(idx => data[i][idx - 1]);
    copiedData.push(row);
  }

  // 新しいスプレッドシートを作成
  // 参加予定日を取得し、ファイル名を設定
  const date = copiedData[1][10];
  const newFileName = YEAR+"参加者リスト_"+date.replaceAll("/","").substring(4)+"_"+createFileName;
  const newFile = SpreadsheetApp.create(newFileName);
  DriveApp.getFileById(newFile.getId()).moveTo(DriveApp.getFolderById(trgFolderId));

  try {
    // カレッジ別でシートを作成　＆　データ書き込み
    createSheetForDepartment(sourceSheet, newFile, copiedData);
  } catch(e) {
    errorMessage('エラー内容：'+e.message);
    console.log('エラー内容：'+e.message);
    deleteSpreadsheetPermanently(newFile.getId());
  }
}

// カレッジ別でシートを作成
function createSheetForDepartment(sourceSheet, newFile, copiedData) {
  // 元データからAG列(33)「接触時志望学科」を取得
  const activeSheet = sourceSheet.getActiveSheet();
  const lastRow = activeSheet.getLastRow();
  const columnRange = activeSheet.getRange(1, 33, lastRow-1);
  const values = columnRange.getValues();

  for(var department of departments) {
    // カレッジのシートを作成
    const newSheet = newFile.insertSheet();
    newSheet.setName(department);  // sheet名をカレッジ名に変更    
    Logger.log("シートの追加: " + newSheet.getName());

    // データの書き込み
    findValueInColumn(values, newSheet, department, copiedData, false);
    
    // 列幅のリサイズ
    newSheet.autoResizeColumns(1, copiedData[0].length);
    // 日本語はリサイズがうまくいかないので調整
    for (var i=1; i<=copiedData[0].length; i++) {
      newSheet.setColumnWidth(i, newSheet.getColumnWidth(i)*2.2);
    }
  }

  // 新しいスプレッドの’シート１’を削除
  newFile.deleteSheet(newFile.getSheetByName('シート1'));
}

// AG列「接触時志望学科」から対象カレッジを検索
function findValueInColumn(values, targetSheet, department, copiedData, protectF) {
  // 1行目に項目を書き込む
  var tmpRanges = targetSheet.getRange(1, 1, 1, copiedData[0].length);
  tmpRanges.setValues([copiedData[0]]);
  tmpRanges.setBackground("#0000cd");
  tmpRanges.setFontColor("#ffffff");
  //　接触メモ項目追加
  tmpRanges = targetSheet.getRange(1, copiedData[0].length + 1, 1, 1);
  tmpRanges.setValues([["接触メモ"]]);
  tmpRanges.setBackground("#008000");
  tmpRanges.setFontColor("#ffffff");

  // 検索 & 書き込み
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].includes(department)) {
      const lastRow = targetSheet.getLastRow();
      targetSheet.getRange(lastRow+1, 1, 1, copiedData[i].length).setValues([copiedData[i]]);
      Logger.log(copiedData[i]);
    }
  }

  // M列「プログラム」でソート
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow()-1, copiedData[0].length).sort(13);
  }
  // セルをプロテクト
  if (protectF == true) {
    const protections = targetSheet.getRange(1, 1, targetSheet.getLastRow(), copiedData[0].length).protect();
    //保護したシートで編集可能なユーザーを取得
    let userList = protections.getEditors();
    //オーナーのみ編集可能にするため、編集ユーザーをすべて削除
    //オーナーの編集権限は削除できないため、オーナーのみ編集可能に
    protections.removeEditors(userList);
  }
}

// ファイル削除
function deleteSpreadsheetPermanently(removeFileId) {
  const token = ScriptApp.getOAuthToken();

  UrlFetchApp.fetch(
    `https://www.googleapis.com/drive/v3/files/${removeFileId}?supportsAllDrives=true`,
    {
      method: 'delete',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  Logger.log('一時ファイルを削除しました');
}

// 作成日取得
function getTodayAsYYYYMMDD() {
  const today = new Date();
  const year = today.getFullYear();
  const month = ('0' + (today.getMonth() + 1)).slice(-2); // 月は0始まりなので+1
  const day = ('0' + today.getDate()).slice(-2);
  const yyyymmdd = `${year}${month}${day}`;
  Logger.log(yyyymmdd);
  return yyyymmdd;
}

// 作成日取得
function getTodayAsMMDD() {
  const today = new Date();
  const month = ('0' + (today.getMonth() + 1)).slice(-2); // 月は0始まりなので+1
  const day = ('0' + today.getDate()).slice(-2);
  const mmdd = `${month}${day}`;
  Logger.log(mmdd);
  return mmdd;
}