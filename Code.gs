// ============================================
// 工廠站別打卡系統 - Google Apps Script 程式碼
// 檔案名稱：Code.gs
// ============================================

// ============================================
// 設定區
// ============================================
var SHEET_NAME_BINDING = '員工綁定資料';  // 綁定資料工作表名稱
var SHEET_NAME_RECORD = '打卡記錄';      // 打卡記錄工作表名稱
var SHEET_NAME_STATIONS = '站別清單';    // 站別清單工作表名稱

// ============================================
// 主要函式：處理網頁請求
// ============================================
function doGet(e) {
  var params = e.parameter;
  
  // 如果有action參數，處理API請求
  if (params.action) {
    return handleAction(params);
  }
  
  // 否則返回前端網頁
  var station = params.station || '';
  var template = HtmlService.createTemplateFromFile('index');
  template.station = station;
  
  return template.evaluate()
    .setTitle('工廠打卡系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================
// 處理API請求
// ============================================
function handleAction(params) {
  var action = params.action;
  var result = {};
  
  try {
    if (action === 'checkDevice') {
      // 檢查裝置是否已綁定
      result = checkDeviceBinding(params.deviceId);
      
    } else if (action === 'bindDevice') {
      // 綁定裝置
      result = bindDevice(params.deviceId, params.employeeId, params.employeeName);
      
    } else if (action === 'checkin') {
      // 打卡
      result = recordCheckin(params.deviceId, params.station);
      
    } else {
      result = {
        success: false,
        message: '未知的操作'
      };
    }
  } catch (error) {
    result = {
      success: false,
      message: '系統錯誤：' + error.toString()
    };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 檢查裝置是否已綁定
// ============================================
function checkDeviceBinding(deviceId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_BINDING);
  
  if (!sheet) {
    return {
      success: false,
      message: '找不到綁定資料工作表'
    };
  }
  
  var data = sheet.getDataRange().getValues();
  
  // 從第2列開始搜尋（第1列是標題）
  for (var i = 1; i < data.length; i++) {
    var rowDeviceId = data[i][2]; // C欄：裝置ID
    
    if (rowDeviceId === deviceId) {
      var employeeId = data[i][0];   // A欄：工號
      var employeeName = data[i][1]; // B欄：姓名
      var status = data[i][4];        // E欄：狀態
      
      if (status === '停用') {
        return {
          success: false,
          isBound: false,
          message: '此裝置已被停用，請聯繫管理員'
        };
      }
      
      return {
        success: true,
        isBound: true,
        employeeId: employeeId,
        employeeName: employeeName
      };
    }
  }
  
  // 未找到綁定記錄
  return {
    success: true,
    isBound: false
  };
}

// ============================================
// 綁定裝置
// ============================================
function bindDevice(deviceId, employeeId, employeeName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_BINDING);
  
  if (!sheet) {
    return {
      success: false,
      message: '找不到綁定資料工作表'
    };
  }
  
  // 檢查工號是否已被其他裝置綁定
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === employeeId && data[i][4] === '啟用') {
      return {
        success: false,
        message: '此工號已綁定其他裝置，請聯繫管理員解除綁定'
      };
    }
  }
  
  // 檢查此裝置是否已綁定其他工號
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === deviceId && data[i][4] === '啟用') {
      return {
        success: false,
        message: '此裝置已綁定工號：' + data[i][0]
      };
    }
  }
  
  // 新增綁定記錄
  var now = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  sheet.appendRow([
    employeeId,
    employeeName,
    deviceId,
    timestamp,
    '啟用'
  ]);
  
  return {
    success: true,
    message: '綁定成功！',
    employeeId: employeeId,
    employeeName: employeeName
  };
}

// ============================================
// 記錄打卡
// ============================================
function recordCheckin(deviceId, station) {
  // 先檢查裝置綁定
  var bindingResult = checkDeviceBinding(deviceId);
  
  if (!bindingResult.success || !bindingResult.isBound) {
    return {
      success: false,
      message: '裝置未綁定，請先完成綁定'
    };
  }
  
  var employeeId = bindingResult.employeeId;
  var employeeName = bindingResult.employeeName;
  
  // 寫入打卡記錄
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_RECORD);
  
  if (!sheet) {
    return {
      success: false,
      message: '找不到打卡記錄工作表'
    };
  }
  
  var now = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  sheet.appendRow([
    timestamp,
    employeeId,
    employeeName,
    station,
    deviceId
  ]);
  
  return {
    success: true,
    message: '打卡成功！',
    employeeId: employeeId,
    employeeName: employeeName,
    station: station,
    time: timestamp
  };
}

// ============================================
// 工具函式：取得站別清單
// ============================================
function getStationList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_STATIONS);
  
  if (!sheet) {
    return ['包裝', 'B區', 'A區', 'A11', 'C區', '發泡', '柴爐', '拉車', '裁切', 'T85', 'T32', 'T55', '倉庫', '其它'];
  }
  
  var data = sheet.getDataRange().getValues();
  var stations = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      stations.push(data[i][0]);
    }
  }
  
  return stations;
}

// ============================================
// 管理函式：解除綁定（可從試算表直接執行）
// ============================================
function unbindDevice(employeeId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_BINDING);
  
  if (!sheet) {
    Logger.log('找不到綁定資料工作表');
    return false;
  }
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === employeeId) {
      // 將狀態改為「停用」
      sheet.getRange(i + 1, 5).setValue('停用');
      Logger.log('已解除工號 ' + employeeId + ' 的綁定');
      return true;
    }
  }
  
  Logger.log('找不到工號 ' + employeeId);
  return false;
}

// ============================================
// 管理函式：查看今日打卡統計
// ============================================
function getTodayStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_RECORD);
  
  if (!sheet) {
    Logger.log('找不到打卡記錄工作表');
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var stats = {};
  
  for (var i = 1; i < data.length; i++) {
    var timestamp = data[i][0];
    var dateStr = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    if (dateStr === today) {
      var station = data[i][3];
      stats[station] = (stats[station] || 0) + 1;
    }
  }
  
  Logger.log('今日打卡統計：');
  for (var station in stats) {
    Logger.log(station + ': ' + stats[station] + ' 次');
  }
  
  return stats;
}
