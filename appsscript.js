/**
 * Paid Leave Management Tool for Google Sheets
 * 有給管理ツール - Google Apps Script
 * 
 * Automatically calculates remaining days, manages grant dates,
 * and tracks expiration using FIFO consumption.
 * 
 * Sheet Structure:
 * Columns A-D: Transaction records (Date, Delta, Running Total, Memo)
 * Column E: Separator (blank)
 * Columns F-K: Grant summary (Grant Date, Days Granted, Used, Remaining, Expiry, Status)
 */

// ============================================================================
// Constants
// ============================================================================

var CONSTANTS = {
  // Sheet structure
  HEADER_ROW: 1,
  DATA_START_ROW: 2,
  
  // Transaction columns (A-D)
  COL_DATE: 1,
  COL_DELTA: 2,
  COL_RUNNING_TOTAL: 3,
  COL_MEMO: 4,
  
  // Separator
  COL_SEPARATOR: 5,
  
  // Summary columns (F-K)
  COL_GRANT_DATE: 6,
  COL_GRANT_DAYS: 7,
  COL_USED: 8,
  COL_REMAINING: 9,
  COL_EXPIRY: 10,
  COL_STATUS: 11,
  
  // Business rules
  EXPIRY_YEARS: 2,
  WARNING_DAYS: 30,
  
  // Status values
  STATUS_VALID: '有効',
  STATUS_WARNING: '期限間近',
  STATUS_USED: '使用済',
  STATUS_EXPIRED: '失効済'
};

// ============================================================================
// Event Handlers
// ============================================================================

/**
 * Triggered when spreadsheet is opened - adds custom menu
 * スプレッドシートが開かれたときにカスタムメニューを追加
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('有給管理')
    .addItem('再計算', 'recalculateAll')
    .addToUi();
}

/**
 * Triggered when cell is edited - processes transactions automatically
 * セルが編集されたときに自動的に処理を実行
 * 
 * @param {Object} e - Edit event object from Google Sheets
 */
function onEdit(e) {
  try {
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
    
    // Skip header row
    if (row < CONSTANTS.DATA_START_ROW) {
      return;
    }
    
    // Only process edits to date (column A) or delta (column B) columns
    if (col !== CONSTANTS.COL_DATE && col !== CONSTANTS.COL_DELTA) {
      return;
    }
    
    autoProcess();
    
  } catch (error) {
    console.error('Error in onEdit:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}

// ============================================================================
// Main Processing Functions
// ============================================================================

/**
 * Main processing function - calculates running totals and updates grant summary
 * メイン処理関数 - 残日数を計算し、付与日別管理を更新
 */
function autoProcess() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow < CONSTANTS.DATA_START_ROW) {
    return;
  }
  
  var transactions = readTransactions(sheet, lastRow);
  var runningTotals = calculateRunningTotals(transactions);
  
  writeRunningTotals(sheet, runningTotals);
  
  var grants = processGrants(transactions);
  var grantSummaries = buildGrantSummaries(grants);
  
  writeGrantSummaries(sheet, grantSummaries);
}

/**
 * Manual recalculation triggered from menu
 * メニューから実行される手動再計算
 */
function recalculateAll() {
  try {
    autoProcess();
    SpreadsheetApp.getUi().alert('再計算が完了しました');
  } catch (error) {
    console.error('Error in recalculateAll:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}

// ============================================================================
// Data Reading Functions
// ============================================================================

/**
 * Reads transaction data from columns A-D
 * A-D列から増減記録データを読み込む
 * 
 * @param {Sheet} sheet - Active spreadsheet sheet
 * @param {number} lastRow - Last row with data
 * @returns {Array<Object>} Array of transaction objects
 */
function readTransactions(sheet, lastRow) {
  var numRows = lastRow - CONSTANTS.HEADER_ROW;
  var data = sheet.getRange(CONSTANTS.DATA_START_ROW, CONSTANTS.COL_DATE, numRows, 4).getValues();
  
  var transactions = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var date = row[0];
    var delta = row[1];
    var memo = row[3];
    
    // Validate data
    if (date instanceof Date && typeof delta === 'number' && !isNaN(delta)) {
      transactions.push({
        date: date,
        delta: delta,
        memo: memo,
        rowIndex: i
      });
    }
  }
  
  return transactions;
}

// ============================================================================
// Calculation Functions
// ============================================================================

/**
 * Calculates running totals for each transaction row
 * 各行の残日数を計算
 * 
 * @param {Array<Object>} transactions - Array of transaction objects
 * @returns {Array<number>} Array of running totals (null for invalid rows)
 */
function calculateRunningTotals(transactions) {
  var runningTotal = 0;
  var totals = [];
  
  for (var i = 0; i < transactions.length; i++) {
    runningTotal += transactions[i].delta;
    totals.push(runningTotal);
  }
  
  return totals;
}

/**
 * Processes grants and usage with FIFO consumption
 * FIFO方式で付与日と利用を処理
 * 
 * @param {Array<Object>} transactions - Array of transaction objects
 * @returns {Array<Object>} Array of grant objects with usage tracking
 */
function processGrants(transactions) {
  // Collect all grants (positive deltas)
  var grants = [];
  for (var i = 0; i < transactions.length; i++) {
    var tx = transactions[i];
    if (tx.delta > 0) {
      grants.push({
        date: tx.date,
        days: tx.delta,
        used: 0
      });
    }
  }
  
  // Sort grants by date (oldest first for FIFO)
  grants.sort(function(a, b) {
    return a.date - b.date;
  });
  
  // Process usage (negative deltas) with FIFO consumption
  for (var i = 0; i < transactions.length; i++) {
    var tx = transactions[i];
    
    if (tx.delta < 0) {
      var useDays = Math.abs(tx.delta);
      
      for (var j = 0; j < grants.length && useDays > 0; j++) {
        // Only consume from grants that existed before the usage date
        if (grants[j].date <= tx.date) {
          var remaining = grants[j].days - grants[j].used;
          
          if (remaining > 0) {
            var consume = Math.min(remaining, useDays);
            grants[j].used += consume;
            useDays -= consume;
          }
        }
      }
    }
  }
  
  return grants;
}

/**
 * Builds grant summary data with status calculation
 * ステータスを含む付与日別管理データを構築
 * 
 * @param {Array<Object>} grants - Array of grant objects
 * @returns {Array<Array>} 2D array for writing to sheet
 */
function buildGrantSummaries(grants) {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var summaries = [];
  
  for (var i = 0; i < grants.length; i++) {
    var grant = grants[i];
    var remaining = grant.days - grant.used;
    
    // Calculate expiry date (grant date + 2 years)
    var expiryDate = new Date(grant.date);
    expiryDate.setFullYear(expiryDate.getFullYear() + CONSTANTS.EXPIRY_YEARS);
    
    // Determine status
    var status = calculateStatus(remaining, expiryDate, today);
    
    summaries.push([
      grant.date,      // F: Grant date
      grant.days,      // G: Days granted
      grant.used,      // H: Used
      remaining,       // I: Remaining
      expiryDate,      // J: Expiry date
      status           // K: Status
    ]);
  }
  
  return summaries;
}

/**
 * Calculates the status of a grant based on remaining days and expiry
 * 残日数と失効日に基づいてステータスを計算
 * 
 * @param {number} remaining - Remaining days
 * @param {Date} expiryDate - Expiry date
 * @param {Date} today - Today's date
 * @returns {string} Status string
 */
function calculateStatus(remaining, expiryDate, today) {
  if (remaining <= 0) {
    return CONSTANTS.STATUS_USED;
  }
  
  if (expiryDate <= today) {
    return CONSTANTS.STATUS_EXPIRED;
  }
  
  var warningDate = new Date(today);
  warningDate.setDate(warningDate.getDate() + CONSTANTS.WARNING_DAYS);
  
  if (expiryDate <= warningDate) {
    return CONSTANTS.STATUS_WARNING;
  }
  
  return CONSTANTS.STATUS_VALID;
}

// ============================================================================
// Data Writing Functions
// ============================================================================

/**
 * Writes running totals to column C
 * 残日数をC列に書き込む
 * 
 * @param {Sheet} sheet - Active spreadsheet sheet
 * @param {Array<number>} totals - Array of running totals
 */
function writeRunningTotals(sheet, totals) {
  if (totals.length === 0) {
    return;
  }
  
  var values = [];
  for (var i = 0; i < totals.length; i++) {
    values.push([totals[i]]);
  }
  
  sheet.getRange(CONSTANTS.DATA_START_ROW, CONSTANTS.COL_RUNNING_TOTAL, values.length, 1).setValues(values);
}

/**
 * Writes grant summaries to columns F-K
 * 付与日別管理をF-K列に書き込む
 * 
 * @param {Sheet} sheet - Active spreadsheet sheet
 * @param {Array<Array>} summaries - 2D array of grant summaries
 */
function writeGrantSummaries(sheet, summaries) {
  var lastRow = sheet.getLastRow();
  
  // Clear existing summary data
  if (lastRow > CONSTANTS.HEADER_ROW) {
    var numRows = lastRow - CONSTANTS.HEADER_ROW;
    sheet.getRange(CONSTANTS.DATA_START_ROW, CONSTANTS.COL_GRANT_DATE, numRows, 6).clearContent();
  }
  
  // Write new data
  if (summaries.length > 0) {
    sheet.getRange(CONSTANTS.DATA_START_ROW, CONSTANTS.COL_GRANT_DATE, summaries.length, 6).setValues(summaries);
  }
}
