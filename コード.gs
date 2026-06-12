/**
 * =================================================================
 * 利用者用Chromebook等 貸出管理システム・コアスクリプト（予備機集計修正版）
 * =================================================================
 */

function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("ログ");
  var deviceSheet = ss.getSheetByName("端末リスト");
  var studentSheet = ss.getSheetByName("生徒一覧"); 
  var activeSheet = ss.getSheetByName("貸出中");

  if (!e || !e.namedValues) return;

  var mode = (e.namedValues['処理モード'] && e.namedValues['処理モード'][0]) ? e.namedValues['処理モード'][0].trim() : "";
  var data1 = (e.namedValues['データ1'] && e.namedValues['データ1'][0]) ? e.namedValues['データ1'][0].trim() : "";
  var data2 = (e.namedValues['データ2'] && e.namedValues['データ2'][0]) ? e.namedValues['データ2'][0].trim() : "";
  var timestamp = (e.namedValues['タイムスタンプ'] && e.namedValues['タイムスタンプ'][0]) ? e.namedValues['タイムスタンプ'][0] : new Date();

  var targetDevice = "";
  var targetStudentId = "";
  var targetStudentName = "";
  var todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");

  var deviceLastRow = deviceSheet.getLastRow();
  var deviceRange = deviceLastRow >= 2 ? deviceSheet.getRange(2, 1, deviceLastRow - 1, 8) : null;
  var deviceData = deviceRange ? deviceRange.getValues() : [];

  var studentLastRow = studentSheet.getLastRow();
  var studentRange = studentLastRow >= 2 ? studentSheet.getRange(2, 1, studentLastRow - 1, 7) : null;
  var studentData = studentRange ? studentRange.getValues() : [];

  var logModeText = mode; 

  try {
    if (mode.indexOf("貸出") !== -1) {
      logModeText = "利用者へ貸出";
      targetStudentId = data1;
      targetDevice = data2;
      var originalOwner = ""; 

      if (deviceRange) {
        for (var i = 0; i < deviceData.length; i++) {
          if (!deviceData[i] || deviceData[i][0] === undefined) continue;
          var colA = String(deviceData[i][0]).trim();
          var colB = String(deviceData[i][1]).trim();
          if (colA === targetDevice || colB === targetDevice) {
            var rowNum = i + 2;
            var currentHValue = String(deviceData[i][7]).trim(); 
            if (currentHValue === "") {
              originalOwner = "通常枠";
            } else {
              originalOwner = currentHValue;
            }
            deviceSheet.getRange(rowNum, 8).setValue(targetStudentId);
            break;
          }
        }
      }

      activeSheet.appendRow([targetDevice, originalOwner, targetStudentId, todayStr]);

      for (var s = 0; s < studentData.length; s++) {
        if (studentData[s] && String(studentData[s][0]).trim() === targetStudentId) {
          targetStudentName = studentData[s][6];
          break;
        }
      }
    }
    else if (mode.indexOf("返却") !== -1) {
      logModeText = "利用者から返却";
      targetDevice = data1; 
      var restoredAssignment = ""; 
      var rentDate = ""; 

      var activeLastRow = activeSheet.getLastRow();
      if (activeLastRow >= 2) {
        var activeData = activeSheet.getRange(2, 1, activeLastRow - 1, 4).getValues();
        for (var a = activeData.length - 1; a >= 0; a--) { 
          if (!activeData[a] || activeData[a][0] === undefined || activeData[a][0] === null || activeData[a][0] === "") {
            continue; 
          }
          if (String(activeData[a][0]).trim() === targetDevice) {
            var savedOwnerInfo = String(activeData[a][1]).trim(); 
            targetStudentId = String(activeData[a][2]).trim();     
            if (savedOwnerInfo === "通常枠") {
              restoredAssignment = "";
            } else {
              restoredAssignment = savedOwnerInfo;
            }
            var origDate = activeData[a][3];
            if (origDate instanceof Date) {
              rentDate = Utilities.formatDate(origDate, "JST", "yyyy/MM/dd");
            } else {
              rentDate = String(origDate).trim();
            }
            activeSheet.deleteRow(a + 2); 
            break;
          }
        }
      }

      if (deviceRange) {
        for (var i = 0; i < deviceData.length; i++) {
          if (!deviceData[i] || deviceData[i][0] === undefined) continue;
          var colA = String(deviceData[i][0]).trim();
          var colB = String(deviceData[i][1]).trim();
          if (colA === targetDevice || colB === targetDevice) {
            var rowNum = i + 2;
            deviceSheet.getRange(rowNum, 8).setValue(restoredAssignment); 
            break;
          }
        }
      }

      if (targetStudentId) {
        for (var s = 0; s < studentData.length; s++) {
          if (studentData[s] && String(studentData[s][0]).trim() === targetStudentId) {
            targetStudentName = studentData[s][6];
            break;
          }
        }
      }
      if (rentDate) {
        logModeText = "利用者から返却 (貸出日: " + rentDate + ")";
      }
    }
    else if (mode.indexOf("新規") !== -1 || mode.indexOf("登録") !== -1) {
      logModeText = "端末新規登録";
      targetDevice = data1;
      var nextId = 1;
      if (deviceLastRow >= 2) {
        var currentMax = deviceSheet.getRange(2, 1).getValue();
        if (!isNaN(currentMax)) nextId = Number(currentMax) + 1;
      }

      var deviceType = "不明な端末";
      if (targetDevice.indexOf("R7HC") === 0) deviceType = "Chromebook本体";
      else if (targetDevice.indexOf("R7AD") === 0) deviceType = "ACアダプタ";
      else if (/^\d{8}$/.test(targetDevice)) deviceType = "モバイルバッテリー";

      deviceSheet.insertRowBefore(2);
      deviceSheet.getRange(2, 1, 1, 8).setValues([[nextId, targetDevice, deviceType, todayStr, "", "配備予定", "正常", ""]]);
    }
    else if (mode.indexOf("解除") !== -1 || mode.indexOf("削除") !== -1) {
      logModeText = "端末登録解除";
      targetDevice = data1;
      if (deviceRange) {
        for (var i = 0; i < deviceData.length; i++) {
          if (!deviceData[i] || deviceData[i][0] === undefined) continue;
          var colA = String(deviceData[i][0]).trim();
          var colB = String(deviceData[i][1]).trim();
          if (colA === targetDevice || colB === targetDevice) {
            var rowNum = i + 2;
            deviceSheet.getRange(rowNum, 5).setValue(todayStr); 
            deviceSheet.getRange(rowNum, 6).setValue("回収済"); 
            deviceSheet.getRange(rowNum, 8).setValue("");
            break;
          }
        }
      }
    }
  } catch (error) {
    logModeText = "⚠️システムエラー発生: " + error.message;
  }
  logSheet.appendRow([timestamp, logModeText, targetDevice, targetStudentId, targetStudentName]);
}

// ② フロント画面用 在庫数・氏名取得API
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var callback = e.parameter.callback;
  var action = e.parameter.action;
  
  // 1. 利用者IDからリアルタイムに氏名を検索（変更なし）
  if (action === "getName") {
    var id = e.parameter.id ? String(e.parameter.id).trim() : "";
    var studentSheet = ss.getSheetByName("生徒一覧");
    var lastRow = studentSheet.getLastRow();
    var name = "⚠️未登録のIDです";
    
    if (lastRow >= 2) {
      var data = studentSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === id) {
          name = data[i][6];
          break;
        }
      }
    }
    var jsonName = JSON.stringify({ name: name });
    return ContentService.createTextOutput(callback + '(' + jsonName + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  // 2. 📊 【本来の正しい在庫数集計ロジック】
  var deviceSheet = ss.getSheetByName("端末リスト");
  var deviceLastRow = deviceSheet.getLastRow();
  
  // A列(ID)〜H列(現在の使用者ID) までの8列分をすべて取得
  var deviceData = deviceLastRow >= 2 ? deviceSheet.getRange(2, 1, deviceLastRow - 1, 8).getValues() : []; 

  var pc99Count = 0;
  var batteryCount = 0;

  // deviceData[i] のインデックス：
  // [1] = B列（端末番号）, [2] = C列（端末種別）, [5] = F列（状態）, [7] = H列（現在の使用者ID）
  for (var i = 0; i < deviceData.length; i++) {
    if (!deviceData[i]) continue;

    var type = String(deviceData[i][2]).trim();      // C列：端末種別
    var status = String(deviceData[i][5]).trim();    // F列：状態
    var borrowerId = String(deviceData[i][7]).trim(); // H列：現在の使用者ID

    // すでに回収・廃棄済みのものは在庫から除外
    if (status === "回収済") continue; 

    // 💻 Chromebook本体のカウントルール
    if (type === "Chromebook本体") {
      // H列（使用者ID）が「99」から始まる架空の生徒IDであれば、予備機在庫としてカウント
      if (borrowerId.indexOf("99") === 0) {
        pc99Count++;
      }
    } 
    // 🔋 モバイルバッテリーのカウントルール
    else if (type === "モバイルバッテリー") {
      // H列（使用者ID）が「空欄」であれば、誰にも貸し出されていない在庫としてカウント
      if (borrowerId === "") {
        batteryCount++;
      }
    }
  }
  
  var jsonOut = JSON.stringify({ pc: pc99Count, battery: batteryCount });
  
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + jsonOut + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService.createTextOutput(jsonOut).setMimeType(ContentService.MimeType.JSON);
  }
}
