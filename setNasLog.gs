function onOpen(){
  SpreadsheetApp.getActiveSpreadsheet().addMenu('NASログ出力', [{name:'NASログ出力', functionName:'setAronasLogs'}]);
}
function setAronasLogs(){
  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  getNasInfo_(outputSheet);
}
/**
 * Obtain the date to be processed.
 * @param none.
 * @return {Array.Date} Array of target dates.
 */
function getTargetDateList_(){
  const today = new Date();
  // Obtains the day of the week of the execution date.
  const todaysDay = today.getDay();
  let yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  let targetDate = [today, yesterday];
  if (todaysDay == 1){
      // If the execution day is Monday, information on Friday and Saturday is obtained in addition to the previous day's information.
    for (let i = 1; i < 3; i++){
      let temp = new Date(yesterday);
      temp.setDate(temp.getDate() - i);
      targetDate.push(temp)
    }
  } else {
    // If the execution date is not a Monday, check if the day before the execution date is a holiday.
    const holiday = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('祝日').getDataRange().getValues().filter(x => new Date(x[0]).getTime()).map(x => Utilities.formatDate(x[0], 'Asia/Tokyo', 'yyyy/MM/dd'));
    const temp = holiday.filter(x => x == Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy/MM/dd'));
    if (temp.length > 0){
      // If today is Tuesday, it should be covered through last Friday. Otherwise, it covers the day before yesterday.
      if (todaysDay == 2){
        for (let i = 1; i < 4; i++){
          let temp = new Date(yesterday);
          temp.setDate(temp.getDate() - i);
          targetDate.push(temp)
        }
      } else {
        let dayBeforeYesterday = new Date(yesterday);
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
function getNasInfo_(outputSheet){
  // Obtain a list of dates to be processed.
  const targetDate = getTargetDateList_();
  // Obtain the line numbers to be output from the date and store them in an array.
  const outputRowIdx = targetDate.map(x => getTargetDateIdx_(outputSheet, 0, x));
  const targetDateString = targetDate.map(x => Utilities.formatDate(x, 'Asia/Tokyo', 'yyyy-MM-dd'));
  const nasLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('wk_nas').getDataRange().getValues();
  // Extract logs for dates to be processed.
  const target = targetDateString.map(x => nasLog.filter(log => new RegExp(x).test(log)));
  outputRowIdx.forEach((x, idx) => {
    const outputRow = x + 1;
    if (outputRow){
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
function getOutputRangesNas_(outputSheet, log, outputRow){
  const initVar = init_();
  const warningString = new RegExp('^Warning');
  const errorString = new RegExp('^Error');
  const hbs = new RegExp('Hybrid Backup Sync');
  const hbsInfo = log.filter(x => hbs.test(x));
  const startIdx = 0;
  const endIdx = 1;
  const jobNameIdx = 2;
  const dateIdx = 3;  
  const hbsStartEndTimeList = initVar.nasJobNameList.map(jobName => {
    let startEnd = [null, null, null, null];
    startEnd[jobNameIdx] = jobName;
    const log = hbsInfo.filter(x => new RegExp(jobName).test(x));
    startEnd[endIdx] = log.map(x => x[0].match(/(?<=^Information,\d{4}-\d{2}-\d{2},)\d{2}:\d{2}:\d{2}(?=.*Finished)/g)).filter(x => x);
    startEnd[startIdx] = log.map(x => x[0].match(/(?<=^Information,\d{4}-\d{2}-\d{2},)\d{2}:\d{2}:\d{2}(?=.*Started)/g)).filter(x => x);
    startEnd[dateIdx] = log.map(x => x[0].match(/(?<=^Information,)\d{4}-\d{2}-\d{2}(?=.*Started)/g)).filter(x => x);
    return startEnd;
  });
  const outputTarget = hbsStartEndTimeList.filter(x => x[dateIdx].length > 0);
  outputTarget.forEach(startEnd =>{
    // Get the columns to output.
    const outputTargetColNum = getColIdx_(outputSheet, 1, startEnd[jobNameIdx]) + 1;
    // Jobs starting before 24:00 will be output on the next date.
    const outputTargetRow = initVar.nasYesterdayStartJobNameList.indexOf(startEnd[jobNameIdx]) > -1 ? outputRow + 1 : outputRow;
    if (startEnd[startIdx].length > 0){
      outputSheet.getRange(outputTargetRow, outputTargetColNum + 1).setValue(startEnd[startIdx]);
    }
    if (startEnd[endIdx].length > 0){
      outputSheet.getRange(outputTargetRow, outputTargetColNum + 2).setValue(startEnd[endIdx]);
    }
    if (startEnd[startIdx].length > 0 && startEnd[endIdx].length > 0){
      outputSheet.getRange(outputTargetRow, outputTargetColNum).setValue('完了');
    }

  });  
  // Warnings and Errors are output to the remarks of today's date.
  const errorAndWarning = log.filter(x => warningString.test(x)|| errorString.test(x));
  if (errorAndWarning.length > 0){
    const outputBikou = errorAndWarning.join('\n');
    const bikouCol = getColIdx_(outputSheet, 2, '備考');
    const temp = outputSheet.getRange(outputRow + 1, bikouCol + 1).getValue();
    const outputBikouString = temp.length > 0 ? temp + '\n' + outputBikou : outputBikou;
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
function getColIdx_(sheet, colRowIdx, colString){
  const target = sheet.getDataRange().getValues()[colRowIdx].map((x, idx) => x == colString ? idx : null).filter(x => x);
  return target[0];
}
/**
 * Returns the index of the column from the column name.
 * @param {Object} Sheet object to be processed.
 * @param {Number} Index of date column, such as 0 for A.
 * @param {Date} Date value.
 * @return {Number} Return the index of the row for that date. such as 9 for the 10th row.
 */
function getTargetDateIdx_(sheet, rowColIdx, rowString){
  const targetDateString = Utilities.formatDate(rowString, 'Asia/Tokyo', 'yyyy-MM-dd');
  const target = sheet.getDataRange().getValues().map((x, idx) => new Date(x[rowColIdx]).getTime() ? Utilities.formatDate(x[rowColIdx], 'Asia/Tokyo', 'yyyy-MM-dd') == targetDateString ? idx : null : null).filter(x => x);
  return target[0];
}
/**
 * Set values for variables used in common functions.
 * @param none.
 * @return {Object} Data commonly needed for each process.
 */
function init_(){
  let initVar = {};
  initVar.nasYesterdayStartJobNameList = ['ARO_backup', 'backupToPotato_Archives', 'backupToPotato_Projects', 'backupToPotato_References'];
  initVar.nasTodayStartJobNameList = ['box_Backup_Datacenter', 'box_Backup_Projects', 'box_Backup_Restricted', 'box_Backup_Shared', 'box_Backup_Stat', 'box_Backup_Trials'];
  initVar.nasJobNameList = initVar.nasTodayStartJobNameList.concat(initVar.nasYesterdayStartJobNameList);
  return initVar;
}
