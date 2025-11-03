// ============================================
// 工廠站別打卡系統 - Google Apps Script 程式碼（修正版）
// 檔案名稱：Code.gs
// 版本：v1.1 - 修正參數傳遞問題
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
  // 新增：檢查參數是否存在
  if (!e || !e.parameter) {
    return createErrorPage('系統錯誤', '無法取得請求參數，請確認QR Code是否正確生成');
  }
  
  var params = e.parameter;
  
  // 記錄請求參數（用於除錯）
  console.log('收到請求，參數：' + JSON.stringify(params));
  
  // 如果有action參數，處理API請求
  if (params.action) {
    return handleAction(params);
  }
  
  // 檢查是否有station參數
  if (!params.station) {
    return createErrorPage('參數錯誤', '未指定站別，請確認QR Code是否包含站別資訊');
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
// 新增：錯誤頁面生成函式
// ============================================
function createErrorPage(title, message) {
  var html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>錯誤</title>
      <style>
        body {
          font-family: 'Microsoft JhengHei', Arial, sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
          display: flex;
          justify-content: center;
          align-items: center;
          margin: 0;
          padding: 20px;
        }
        .error-container {
          background: white;
          border-radius: 20px;
          padding: 40px;
          max-width: 500px;
          text-align: center;
          box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }
        .error-icon {
          font-size: 60px;
          margin-bottom: 20px;
        }
        h1 {
          color: #dc3545;
          margin-bottom: 20px;
        }
        p {
          color: #666;
          line-height: 1.6;
          margin-bottom: 30px;
        }
        .help-box {
          background: #f8f9fa;
          padding: 20px;
          border-radius: 10px;
          text-align: left;
        }
        .help-box h3 {
          color: #333;
          margin-bottom: 10px;
        }
        .help-box ol {
          margin-left: 20px;
          color: #666;
        }
      </style>
    </head>
    <body>
      <div class="error-container">
        <div class="error-icon">❌</div>
        <h1>${title}</h1>
        <p>${message}</p>
        
        <div class="help-box">
          <h3>🔧 解決方法</h3>
          <ol>
            <li>確認QR Code格式正確</li>
            <li>確認網址包含站別參數</li>
            <li>重新生成QR Code</li>
            <li>聯繫系統管理員</li>
          </ol>
        </div>
      </div>
    </body>
    </html>
  `);
  
  return html.setTitle('錯誤');
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
      if (!params.deviceId) {
        throw new Error('缺少裝置ID參數');
      }
      result = checkDeviceBinding(params.deviceId);
      
    } else if (action === 'bindDevice') {
      // 綁定裝置
      if (!params.deviceId || !params.employeeId || !params.employeeName) {
        throw new Error('綁定參數不完整');
      }
      result = bindDevice(params.deviceId, params.employeeId, params.employeeName);
      
    } else if (action === 'checkin') {
      // 打卡
      if (!params.deviceId || !params.station) {
        throw new Error('打卡參數不完整');
      }
      result = recordCheckin(params.deviceId, params.station);
      
    } else {
      result = {
        success: false,
        message: '未知的操作：' + action
      };
    }
  } catch (error) {
    console.error('處理請求時發生錯誤：' + error.toString());
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
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME_BINDING);
    
    if (!sheet) {
      return {
        success: false,
        message: '找不到「' + SHEET_NAME_BINDING + '」工作表，請確認工作表名稱是否正確'
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
  } catch (error) {
    console.error('檢查裝置綁定時發生錯誤：' + error.toString());
    return {
      success: false,
      message: '檢查綁定失敗：' + error.toString()
    };
  }
}

// ============================================
// 綁定裝置
// ============================================
function bindDevice(deviceId, employeeId, employeeName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME_BINDING);
    
    if (!sheet) {
      return {
        success: false,
        message: '找不到「' + SHEET_NAME_BINDING + '」工作表'
      };
    }
    
    // 檢查工號是否已被其他裝置綁定
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId && data[i][4] === '啟用') {
        return {
          success: false,
          message: '此工號已綁定其他裝置（裝置ID：' + data[i][2].substring(0, 16) + '...），請聯繫管理員解除綁定'
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
    
    console.log('綁定成功：工號=' + employeeId + ', 姓名=' + employeeName);
    
    return {
      success: true,
      message: '綁定成功！',
      employeeId: employeeId,
      employeeName: employeeName
    };
  } catch (error) {
    console.error('綁定裝置時發生錯誤：' + error.toString());
    return {
      success: false,
      message: '綁定失敗：' + error.toString()
    };
  }
}

// ============================================
// 記錄打卡
// ============================================
function recordCheckin(deviceId, station) {
  try {
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
        message: '找不到「' + SHEET_NAME_RECORD + '」工作表'
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
    
    console.log('打卡成功：工號=' + employeeId + ', 站別=' + station + ', 時間=' + timestamp);
    
    return {
      success: true,
      message: '打卡成功！',
      employeeId: employeeId,
      employeeName: employeeName,
      station: station,
      time: timestamp
    };
  } catch (error) {
    console.error('記錄打卡時發生錯誤：' + error.toString());
    return {
      success: false,
      message: '打卡失敗：' + error.toString()
    };
  }
}

// ============================================
// 工具函式：取得站別清單
// ============================================
function getStationList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME_STATIONS);
    
    if (!sheet) {
      // 如果沒有站別清單，返回預設站別
      return ['包裝', 'B區', 'A區', 'A11', 'C區', '發泡', '柴爐', '拉車', '裁切', 'T85', 'T32', 'T55', '倉庫', '其它'];
    }
    
    var data = sheet.getDataRange().getValues();
    var stations = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        stations.push(data[i][0]);
      }
    }
    
    return stations.length > 0 ? stations : ['包裝', 'B區', 'A區', 'A11', 'C區', '發泡', '柴爐', '拉車', '裁切', 'T85', 'T32', 'T55', '倉庫', '其它'];
  } catch (error) {
    console.error('取得站別清單時發生錯誤：' + error.toString());
    return ['包裝', 'B區', 'A區', 'A11', 'C區', '發泡', '柴爐', '拉車', '裁切', 'T85', 'T32', 'T55', '倉庫', '其它'];
  }
}

// ============================================
// 管理函式：解除綁定（可從試算表直接執行）
// ============================================
function unbindDevice(employeeId) {
  try {
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
  } catch (error) {
    Logger.log('解除綁定時發生錯誤：' + error.toString());
    return false;
  }
}

// ============================================
// 管理函式：查看今日打卡統計
// ============================================
function getTodayStats() {
  try {
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
  } catch (error) {
    Logger.log('查看統計時發生錯誤：' + error.toString());
    return {};
  }
}

// ============================================
// 新增：測試函式（用於除錯）
// ============================================
function testDoGet() {
  // 模擬掃描QR Code的請求
  var testEvent = {
    parameter: {
      station: '包裝'
    }
  };
  
  var result = doGet(testEvent);
  Logger.log('測試結果：' + result.getContent());
}

function testCheckDevice() {
  // 測試檢查裝置綁定
  var testDeviceId = 'TEST_DEVICE_12345';
  var result = checkDeviceBinding(testDeviceId);
  Logger.log('測試結果：' + JSON.stringify(result));
}
