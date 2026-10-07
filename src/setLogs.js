const awsSizeIdx = 2;
// Script property key for the Google Drive folder ID where the S3 backup list (YYYYMMDD.txt) is uploaded.
const AWS_LOG_FOLDER_ID_KEY = 'AWS_LOG_FOLDER_ID';
// True when run by a time-driven trigger (no UI available).
let isTriggerRun_ = false;
// Background color for cells that could not be imported when run by a trigger.
const AWS_ERROR_BACKGROUND = '#ff0000';
// Background color for cells of closed servers on and after the end date.
const AWS_CLOSED_BACKGROUND = '#cccccc';

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('ログ出力', [
    { name: 'NAS', functionName: 'setAronasLogs' },
    { name: 'AWS（手動貼り付け）', functionName: 'setAwsLogFromSheet' },
  ]);
}
/**
 * Set a dummy value for the script property. Run once, then replace it with the actual folder ID.
 * @param none.
 * @return none.
 */
function initScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(AWS_LOG_FOLDER_ID_KEY) === null) {
    properties.setProperty(AWS_LOG_FOLDER_ID_KEY, 'dummy');
  }
}
/**
 * Read today's S3 backup list file from Google Drive.
 * @param {String} Today's date string (yyyyMMdd).
 * @return {Array.String} Lines of the file. null if the file does not exist.
 */
function getAwsLogLines_(todayYYYYMMDD) {
  const folderId = PropertiesService.getScriptProperties().getProperty(
    AWS_LOG_FOLDER_ID_KEY
  );
  const files = DriveApp.getFolderById(folderId).getFilesByName(
    todayYYYYMMDD + '.txt'
  );
  if (!files.hasNext()) {
    return null;
  }
  return files
    .next()
    .getBlob()
    .getDataAsString('UTF-8')
    .split(/\r?\n/)
    .filter(x => x !== '');
}
/**
 * Read the S3 backup list from the file on Google Drive and output it.
 * @param none.
 * @return none.
 */
function setAwsLog() {
  const today = new Date();
  const todayYYYYMMDD = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd');
  const inputLines = getAwsLogLines_(todayYYYYMMDD);
  if (inputLines === null) {
    outputMsg_(
      [todayYYYYMMDD + '.txt'],
      ' がGoogle Driveに見つかりません。wk_awsシートに貼り付けて「AWS（手動貼り付け）」を実行してください'
    );
    markAllAwsCellsAsError_(today);
    appendAwsBikouForToday_(
      today,
      '【AWS】' + todayYYYYMMDD + '.txt がGoogle Driveに見つかりません'
    );
    return;
  }
  outputAwsLog_(inputLines, today, todayYYYYMMDD);
}
/**
 * Entry point for the time-driven trigger. Runs setAwsLog on weekdays only.
 * @param none.
 * @return none.
 */
function setAwsLogByTrigger() {
  const todaysDay = new Date().getDay();
  // Skip Saturday and Sunday.
  if (todaysDay === 0 || todaysDay === 6) {
    return;
  }
  isTriggerRun_ = true;
  setAwsLog();
}
/**
 * Read the S3 backup list pasted into the wk_aws sheet and output it (for when the automatic upload fails).
 * @param none.
 * @return none.
 */
function setAwsLogFromSheet() {
  const inputSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('wk_aws');
  const inputLines = inputSheet
    .getDataRange()
    .getValues()
    .map(x => x[0]);
  const today = new Date();
  const todayYYYYMMDD = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd');
  outputAwsLog_(inputLines, today, todayYYYYMMDD);
}
/**
 * Edit the S3 backup list and output to a spreadsheet.
 * @param {Array.String} Lines of the S3 backup list.
 * @param {Date} Today's date.
 * @param {String} Today's date string (yyyyMMdd).
 * @return none.
 */
function outputAwsLog_(inputLines, today, todayYYYYMMDD) {
  // Only items dated today are eligible.
  // date, time, size, unit, dump name
  const valueTableSplitBySpace = inputLines.map(x => x.split(/\s+/));
  const targetValues = valueTableSplitBySpace.filter(x =>
    new RegExp(todayYYYYMMDD).test(x)
  );
  if (targetValues.length === 0) {
    markAllAwsCellsAsError_(today);
    appendAwsBikouForToday_(
      today,
      '【AWS】' + todayYYYYMMDD + '.txt に今日のバックアップがありません'
    );
    return;
  }
  const dumpNameIdx = 4;
  const serverNameIdx = valueTableSplitBySpace[0].length;
  const removeFileNameFoot = new RegExp('[/|_]' + todayYYYYMMDD + '.dump');
  const valueTableSplitDumpName = valueTableSplitBySpace.map(x =>
    x.concat(x[dumpNameIdx].replace(removeFileNameFoot, ''))
  );
  const outputSheet = getOutputSheet_();
  const outputRow = getTargetDateIdx_(outputSheet, 0, today) + 1;
  const outputValueArray = valueTableSplitDumpName.map(x =>
    x.concat(getColIdx_(outputSheet, 1, x[serverNameIdx]))
  );
  const serverExistenceCheckBackupString = outputValueArray
    .map(x => x[serverNameIdx])
    .flat();
  const serverExistenceCheckBackupOutputSheet =
    getOutputSheetAwsServerNames_(outputSheet);
  const check1 = serverExistenceCheckBackupString.filter(
    x => !serverExistenceCheckBackupOutputSheet.includes(x)
  );
  if (check1.length > 0) {
    outputMsg_(check1, 'の出力列を追加して再実行してください');
    markAllAwsCellsAsError_(today);
    appendAwsBikou_(
      outputSheet,
      outputRow,
      '【AWS】' + check1.join(', ') + ' の出力列がありません'
    );
    return;
  }
  const check2 = check2Aws_(
    serverExistenceCheckBackupOutputSheet,
    serverExistenceCheckBackupString,
    outputSheet
  );
  const outputValueAndCol =
    check2 !== null
      ? [...getAwsOutputValues_(outputValueArray), ...check2.values()]
      : getAwsOutputValues_(outputValueArray);
  outputValueAndCol.forEach(([outputValue, colIdx]) =>
    outputSheet.getRange(outputRow, colIdx + 1).setValue(outputValue)
  );
  clearAwsErrorBackground_(outputSheet, outputRow, outputValueAndCol);
  grayOutClosedServers_(outputSheet);
  if (check2 === null) {
    return;
  }
  if (check2.size > 0) {
    const target = Array.from(check2.keys());
    outputMsg_(target, 'のバックアップを確認してください');
    appendAwsBikou_(
      outputSheet,
      outputRow,
      '【AWS】' + target.join(', ') + ' のバックアップを確認してください'
    );
    markAwsCellsAsError_(
      outputSheet,
      outputRow,
      Array.from(check2.values()).map(([, colIdx]) => colIdx)
    );
  }
}
/**
 * Color the cells that could not be imported. Only when run by a trigger.
 * @param {Object} The object of the sheet to output.
 * @param {Number} Row number of the spreadsheet to output.
 * @param {Array.Number} Indexes of the columns to color, such as 0 for A.
 * @return none.
 */
function markAwsCellsAsError_(outputSheet, outputRow, colIdxList) {
  if (!isTriggerRun_) {
    return;
  }
  const a1Notations = colIdxList.map(colIdx =>
    outputSheet.getRange(outputRow, colIdx + 1).getA1Notation()
  );
  outputSheet.getRangeList(a1Notations).setBackground(AWS_ERROR_BACKGROUND);
}
/**
 * Gray out the cells of closed servers on and after the end date (column B of wk_closed_servers).
 * Servers without an end date are not grayed out.
 * @param {Object} The object of the sheet to output.
 * @return none.
 */
function grayOutClosedServers_(outputSheet) {
  const closedServers = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('wk_closed_servers')
    .getRange('A:B')
    .getValues()
    .filter(([name, endDate]) => name !== '' && endDate instanceof Date);
  if (closedServers.length === 0) {
    return;
  }
  // Date string (yyyy-MM-dd) of each row in column A. null if not a date.
  const rowDateStrings = outputSheet
    .getDataRange()
    .getValues()
    .map(row =>
      row[0] instanceof Date
        ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy-MM-dd')
        : null
    );
  const a1Notations = closedServers
    .map(([name, endDate]) => {
      const colIdx = getColIdx_(outputSheet, 1, name);
      if (colIdx === undefined) {
        return [];
      }
      const endDateString = Utilities.formatDate(
        endDate,
        'Asia/Tokyo',
        'yyyy-MM-dd'
      );
      return rowDateStrings
        .map((dateString, rowIdx) =>
          dateString !== null && dateString >= endDateString
            ? outputSheet.getRange(rowIdx + 1, colIdx + 1).getA1Notation()
            : null
        )
        .filter(x => x !== null);
    })
    .flat();
  if (a1Notations.length === 0) {
    return;
  }
  outputSheet.getRangeList(a1Notations).setBackground(AWS_CLOSED_BACKGROUND);
}
/**
 * Reset the error color of cells where a value has been entered. Other colors are left unchanged.
 * @param {Object} The object of the sheet to output.
 * @param {Number} Row number of the spreadsheet to output.
 * @param {Array} Pairs of [output value, column index].
 * @return none.
 */
function clearAwsErrorBackground_(outputSheet, outputRow, outputValueAndCol) {
  outputValueAndCol
    .filter(([outputValue]) => outputValue !== '')
    .map(([, colIdx]) => outputSheet.getRange(outputRow, colIdx + 1))
    .filter(range => range.getBackground() === AWS_ERROR_BACKGROUND)
    .forEach(range => range.setBackground(null));
}
/**
 * Append a message to the remarks column. Only when run by a trigger.
 * @param {Object} The object of the sheet to output.
 * @param {Number} Row number of the spreadsheet to output.
 * @param {String} Message to append.
 * @return none.
 */
function appendAwsBikou_(outputSheet, outputRow, message) {
  if (!isTriggerRun_) {
    return;
  }
  const bikouCol = getColIdx_(outputSheet, 2, '備考');
  const bikouRange = outputSheet.getRange(outputRow, bikouCol + 1);
  const saveBikouValue = bikouRange.getValue();
  // Do not append the same message twice.
  if (saveBikouValue.includes(message)) {
    return;
  }
  bikouRange.setValue(
    saveBikouValue.length > 0 ? saveBikouValue + '\n' + message : message
  );
}
/**
 * Append a message to the remarks column of today's row. Only when run by a trigger.
 * @param {Date} Today's date.
 * @param {String} Message to append.
 * @return none.
 */
function appendAwsBikouForToday_(today, message) {
  if (!isTriggerRun_) {
    return;
  }
  const outputSheet = getOutputSheet_();
  const outputRow = getTargetDateIdx_(outputSheet, 0, today) + 1;
  appendAwsBikou_(outputSheet, outputRow, message);
}
/**
 * Color today's cells of all AWS servers. Only when run by a trigger.
 * @param {Date} Today's date.
 * @return none.
 */
function markAllAwsCellsAsError_(today) {
  if (!isTriggerRun_) {
    return;
  }
  const outputSheet = getOutputSheet_();
  const outputRow = getTargetDateIdx_(outputSheet, 0, today) + 1;
  const colIdxList = getOutputSheetAwsServerNames_(outputSheet).map(name =>
    getColIdx_(outputSheet, 1, name)
  );
  markAwsCellsAsError_(outputSheet, outputRow, colIdxList);
}
function getAwsOutputValues_(values) {
  const colIdxIdx = values[0].length - 1;
  const result = values.map(x => [x[awsSizeIdx], x[colIdxIdx]]);
  return result;
}
function check2Aws_(
  serverExistenceCheckBackupOutputSheet,
  serverExistenceCheckBackupString,
  outputSheet
) {
  const errorValue = '';
  const check2 = serverExistenceCheckBackupOutputSheet.filter(
    x => !serverExistenceCheckBackupString.includes(x)
  );
  if (check2.length === 0) {
    return null;
  }
  const targetNameMapErrorValueAndColIdx = new Map();
  check2.forEach(name => {
    const colIdx = getColIdx_(outputSheet, 1, name);
    targetNameMapErrorValueAndColIdx.set(name, [errorValue, colIdx]);
  });
  return targetNameMapErrorValueAndColIdx;
}
/**
 * Output of pop-up messages.
 * @param <Array.String>
 * @param <String>
 * @return none.
 */
function outputMsg_(target, msg) {
  const messageString = target.length === 1 ? target : target.join(', ');
  // Pop-ups cannot be displayed when run by a trigger, so write to the execution log instead.
  if (isTriggerRun_) {
    console.log(messageString + msg);
    return;
  }
  Browser.msgBox(messageString + msg);
}
function getOutputSheetAwsServerNames_(outputSheet) {
  // Exclude stopped backups
  const excludeServerNames = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('wk_closed_servers')
    .getRange('A:A')
    .getValues()
    .filter(x => x !== '')
    .flat();
  const awsColStart = getColIdx_(outputSheet, 0, 'AWS') + 1;
  const colCount = outputSheet.getLastColumn() - awsColStart;
  const targetStrings = outputSheet
    .getRange(1, awsColStart + 1, 1, colCount)
    .getValues()[0];
  const colCheck = targetStrings
    .map((x, idx) => (x !== '' ? idx : null))
    .filter(x => x);
  const awsColEnd =
    colCheck.length === 0
      ? outputSheet.getLastColumn()
      : awsColStart + colCheck[0] + 1;
  const serverNames = outputSheet
    .getRange(2, awsColStart, 1, awsColEnd - awsColStart)
    .getValues()[0]
    .filter(x => x !== '');
  const resServerNames =
    excludeServerNames.length > 0
      ? serverNames.filter(x => !excludeServerNames.includes(x))
      : serverNames;
  return resServerNames;
}
function setAronasLogs() {
  const outputSheet = getOutputSheet_();
  getNasInfo_(outputSheet);
}
function getOutputSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}
/**
 * Obtain the date to be processed.
 * @param none.
 * @return {Array.Date} Array of target dates.
 */
function getTargetDateList_() {
  const today = new Date();
  // Obtains the day of the week of the execution date.
  const todaysDay = today.getDay();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const targetDate = [today, yesterday];
  if (todaysDay === 1) {
    // If the execution day is Monday, information on Friday and Saturday is obtained in addition to the previous day's information.
    for (let i = 1; i < 3; i++) {
      const temp = new Date(yesterday);
      temp.setDate(temp.getDate() - i);
      targetDate.push(temp);
    }
  } else {
    // If the execution date is not a Monday, check if the day before the execution date is a holiday.
    const holiday = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('祝日')
      .getDataRange()
      .getValues()
      .filter(x => new Date(x[0]).getTime())
      .map(x => Utilities.formatDate(x[0], 'Asia/Tokyo', 'yyyy/MM/dd'));
    const temp = holiday.filter(
      x => x === Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy/MM/dd')
    );
    if (temp.length > 0) {
      // If today is Tuesday, it should be covered through last Friday. Otherwise, it covers the day before yesterday.
      if (todaysDay === 2) {
        for (let i = 1; i < 4; i++) {
          const temp = new Date(yesterday);
          temp.setDate(temp.getDate() - i);
          targetDate.push(temp);
        }
      } else {
        const dayBeforeYesterday = new Date(yesterday);
        dayBeforeYesterday.setDate(dayBeforeYesterday.getDate() - 1);
        targetDate.push(dayBeforeYesterday);
      }
    }
  }
  return targetDate;
}
/**
 * Edit the NAS logs and output to a spreadsheet.
 * @param {Object} The object of the sheet to output.
 * @return none.
 */
function getNasInfo_(outputSheet) {
  // Obtain a list of dates to be processed.
  const targetDate = getTargetDateList_();
  // Obtain the line numbers to be output from the date and store them in an array.
  const outputRowIdx = targetDate.map(x =>
    getTargetDateIdx_(outputSheet, 0, x)
  );
  const targetDateString = targetDate.map(x =>
    Utilities.formatDate(x, 'Asia/Tokyo', 'yyyy-MM-dd')
  );
  const nasLogSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('wk_nas');
  const nasLogLastRow = nasLogSheet.getLastRow();
  const nasLog = nasLogSheet.getRange(1, 1, nasLogLastRow, 1).getValues();
  // Extract logs for dates to be processed.
  const target = targetDateString.map(x =>
    nasLog.filter(log => new RegExp(x).test(log))
  );
  outputRowIdx.forEach((x, idx) => {
    const outputRow = x + 1;
    if (outputRow) {
      getOutputRangesNas_(outputSheet, target[idx], outputRow);
    }
  });
}
/**
 * Edit the NAS logs and output to a spreadsheet.
 * @param {Object} The object of the sheet to output.
 * @param {Array.String} Log string.
 * @param {Number} Row number of the spreadsheet to output.
 * @return none.
 */
function getOutputRangesNas_(outputSheet, log, outputRow) {
  const initVar = nasInit_();
  const warningString = new RegExp('^Warning');
  const errorString = new RegExp('^Error');
  const hbs = new RegExp('Hybrid Backup Sync');
  const hbsInfo = log.filter(x => hbs.test(x));
  const startIdx = 0;
  const endIdx = 1;
  const jobNameIdx = 2;
  const dateIdx = 3;
  const hbsStartEndTimeList = initVar.nasJobNameList.map(jobName => {
    const startEnd = [null, null, null, null];
    startEnd[jobNameIdx] = jobName;
    const log = hbsInfo.filter(x => new RegExp(jobName).test(x));
    startEnd[endIdx] = log
      .map(x =>
        x[0].match(
          /(?<=^Information,\d{4}-\d{2}-\d{2},)\d{2}:\d{2}:\d{2}(?=.*Finished)/g
        )
      )
      .filter(x => x);
    startEnd[startIdx] = log
      .map(x =>
        x[0].match(
          /(?<=^Information,\d{4}-\d{2}-\d{2},)\d{2}:\d{2}:\d{2}(?=.*Started)/g
        )
      )
      .filter(x => x);
    startEnd[dateIdx] = log
      .map(x => x[0].match(/(?<=^Information,)\d{4}-\d{2}-\d{2}(?=.*Started)/g))
      .filter(x => x);
    if (startEnd[dateIdx] === null || startEnd[dateIdx].length === 0) {
      if (
        jobName === 'box_Backup_Datacenter' ||
        jobName === 'box_Backup_Trials'
      ) {
        startEnd[dateIdx] = log
          .map(x =>
            x[0].match(/(?<=^Information,)\d{4}-\d{2}-\d{2}(?=.*Finished)/g)
          )
          .filter(x => x);
      }
    }
    return startEnd;
  });
  const outputTarget = hbsStartEndTimeList.filter(x => x[dateIdx].length > 0);
  outputTarget.forEach(startEnd => {
    // Get the columns to output.
    const outputTargetColNum =
      getColIdx_(outputSheet, 1, startEnd[jobNameIdx]) + 1;
    // Jobs starting before 24:00 will be output on the next date.
    const outputTargetRow =
      initVar.nasYesterdayStartJobNameList.indexOf(startEnd[jobNameIdx]) > -1
        ? outputRow + 1
        : outputRow;
    if (
      outputSheet.getRange(outputTargetRow, outputTargetColNum).getValue()
        .length === 0
    ) {
      if (startEnd[startIdx].length > 0) {
        outputSheet
          .getRange(outputTargetRow, outputTargetColNum + 1)
          .setValue(startEnd[startIdx]);
      }
      if (startEnd[endIdx].length > 0) {
        outputSheet
          .getRange(outputTargetRow, outputTargetColNum + 2)
          .setValue(startEnd[endIdx]);
      }
      if (startEnd[startIdx].length > 0 && startEnd[endIdx].length > 0) {
        outputSheet
          .getRange(outputTargetRow, outputTargetColNum)
          .setValue('完了');
      }
      if (
        startEnd[jobNameIdx] === 'box_Backup_Datacenter' ||
        startEnd[jobNameIdx] === 'box_Backup_Trials'
      ) {
        if (startEnd[endIdx].length > 0 && startEnd[dateIdx].length > 0) {
          outputSheet
            .getRange(outputTargetRow, outputTargetColNum)
            .setValue('完了');
        }
      }
    }
  });
  // Warnings and Errors are output to the remarks of today's date.
  const errorAndWarning = log.filter(
    x => warningString.test(x) || errorString.test(x)
  );
  if (errorAndWarning.length > 0) {
    const outputBikou = errorAndWarning.join('\n');
    const bikouCol = getColIdx_(outputSheet, 2, '備考');
    const saveBikouValue = outputSheet
      .getRange(outputRow, bikouCol + 1)
      .getValue();
    let temp = saveBikouValue;
    // Remove duplicate values.
    temp = errorAndWarning.reduce(
      (totalValue, currentValue) => totalValue.replace(currentValue, ''),
      saveBikouValue
    );
    // Remove consecutive line breaks
    temp = temp.replace(/(?<=\n)\n/g, '');
    temp = temp.replace(/^\n+/g, '');
    const outputBikouString =
      temp.length > 0 ? temp + '\n' + outputBikou : outputBikou;
    outputSheet.getRange(outputRow, bikouCol + 1).setValue(outputBikouString);
  }
}
/**
 * Returns the index of the column from the column name.
 * @param {Object} Sheet object to be processed.
 * @param {Number} Index of the header row, such as 0 for the first row.
 * @param {String} String of column name.
 * @return {Number} Index of the header column, such as 0 for A.
 */
function getColIdx_(sheet, colRowIdx, colString) {
  const target = sheet
    .getDataRange()
    .getValues()
    [colRowIdx].map((x, idx) => (x === colString ? idx : null))
    .filter(x => x);
  return target[0];
}
/**
 * Returns the index of the column from the column name.
 * @param {Object} Sheet object to be processed.
 * @param {Number} Index of date column, such as 0 for A.
 * @param {Date} Date value.
 * @return {Number} Return the index of the row for that date. such as 9 for the 10th row.
 */
function getTargetDateIdx_(sheet, rowColIdx, rowString) {
  const targetDateString = Utilities.formatDate(
    rowString,
    'Asia/Tokyo',
    'yyyy-MM-dd'
  );
  const target = sheet
    .getDataRange()
    .getValues()
    .map((x, idx) =>
      new Date(x[rowColIdx]).getTime()
        ? Utilities.formatDate(x[rowColIdx], 'Asia/Tokyo', 'yyyy-MM-dd') ==
          targetDateString
          ? idx
          : null
        : null
    )
    .filter(x => x);
  return target[0];
}
/**
 * Set values for variables used in common functions.
 * @param none.
 * @return {Object} Data commonly needed for each process.
 */
function nasInit_() {
  const initVar = {};
  const jobnameSs =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('wk_nas_jobname');
  const bodyStartRow = 2;
  const yesterdayCol = 1;
  const todayCol = 2;
  const nasYesterdayStartJobNameList = jobnameSs
    .getRange(bodyStartRow, yesterdayCol, jobnameSs.getLastRow(), 1)
    .getValues()
    .flat()
    .filter(x => x.length > 0);
  const nasTodayStartJobNameList = jobnameSs
    .getRange(bodyStartRow, todayCol, jobnameSs.getLastRow(), 1)
    .getValues()
    .flat()
    .filter(x => x.length > 0);
  initVar.nasYesterdayStartJobNameList = nasYesterdayStartJobNameList;
  initVar.nasTodayStartJobNameList = nasTodayStartJobNameList;
  initVar.nasJobNameList = initVar.nasTodayStartJobNameList.concat(
    initVar.nasYesterdayStartJobNameList
  );
  return initVar;
}
