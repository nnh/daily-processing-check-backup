function onOpen(){
  SpreadsheetApp.getActiveSpreadsheet().addMenu('ログ出力', [{name:'NAS', functionName:'setAronasLogs'}, {name:'AWS', functionName:'setAwsLog'}]);
}
function setAwsLog(){
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('wk_aws');
  const inputValues = inputSheet.getDataRange().getValues();
  // Only items dated today are eligible.
  const today = new Date();
  const todayYYYYMMDD = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd');
  // date, time, size, unit, dump name
  const valueTableSplitBySpace = inputValues.map(x => x[0].split(/\s+/));
  const targetValues = valueTableSplitBySpace.filter(x => new RegExp(todayYYYYMMDD).test(x));
  if (targetValues.length == 0){
    return;
  }
  const dumpNameIdx = 4;
  const sizeIdx = 2;
  const serverNameIdx = valueTableSplitBySpace[0].length;
  const removeFileNameFoot = new RegExp('[\/|_]' + todayYYYYMMDD + '.dump');
  const valueTableSplitDumpName = valueTableSplitBySpace.map(x => x.concat(x[dumpNameIdx].replace(removeFileNameFoot, '')));
  const outputSheet = getOutputSheet_();
  const outputRow = getTargetDateIdx_(outputSheet, 0, today) + 1;
  const colIdxIdx = valueTableSplitDumpName[0].length; 
  const outputValues = valueTableSplitDumpName.map(x => x.concat(getColIdx_(outputSheet, 1, x[serverNameIdx])));
  const checkOutputCol = outputValues.map(x => !x[colIdxIdx] ? x : null).filter(x => x);
  if (checkOutputCol.length > 0){
    const targetServerName = checkOutputCol.map(x => x[serverNameIdx]).join(',');
    Browser.msgBox(targetServerName + 'の出力列を追加して再実行してください');
    return;
  }
  outputValues.forEach(x => {
    outputSheet.getRange(outputRow, x[colIdxIdx] + 1).setValue(x[sizeIdx]);
  }); 
}
function setAronasLogs(){
  const outputSheet = getOutputSheet_();
  getNasInfo_(outputSheet);
}
function getOutputSheet_(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
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
  const nasLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('wk_nas');
  const nasLogLastRow = nasLogSheet.getLastRow();
  const nasLog = nasLogSheet.getRange(1, 1, nasLogLastRow, 1).getValues();
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
    if (outputSheet.getRange(outputTargetRow, outputTargetColNum).getValue().length == 0){
      if (startEnd[startIdx].length > 0){
        outputSheet.getRange(outputTargetRow, outputTargetColNum + 1).setValue(startEnd[startIdx]);
      }
      if (startEnd[endIdx].length > 0){
        outputSheet.getRange(outputTargetRow, outputTargetColNum + 2).setValue(startEnd[endIdx]);
      }
      if (startEnd[startIdx].length > 0 && startEnd[endIdx].length > 0){
        outputSheet.getRange(outputTargetRow, outputTargetColNum).setValue('完了');
      }
    }
  });  
  // Warnings and Errors are output to the remarks of today's date.
  const errorAndWarning = log.filter(x => warningString.test(x)|| errorString.test(x));
  if (errorAndWarning.length > 0){
    const outputBikou = errorAndWarning.join('\n');
    const bikouCol = getColIdx_(outputSheet, 2, '備考');
    const saveBikouValue = outputSheet.getRange(outputRow, bikouCol + 1).getValue();
    let temp = saveBikouValue;
    // Remove duplicate values.
    temp = errorAndWarning.reduce((totalValue, currentValue) => totalValue.replace(currentValue, ''), saveBikouValue);
    // Remove consecutive line breaks
    temp = temp.replace(/(?<=\n)\n/g, '');
    temp = temp.replace(/^\n+/g, '');
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
function nasInit_(){
  let initVar = {};
  const jobnameSs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('wk_nas_jobname');
  const bodyStartRow = 2;
  const yesterdayCol = 1;
  const todayCol = 2;
  const nasYesterdayStartJobNameList = jobnameSs.getRange(bodyStartRow, yesterdayCol, jobnameSs.getLastRow(), 1).getValues().flat().filter(x => x.length > 0);
  const nasTodayStartJobNameList = jobnameSs.getRange(bodyStartRow, todayCol, jobnameSs.getLastRow(), 1).getValues().flat().filter(x => x.length > 0);
  initVar.nasYesterdayStartJobNameList = nasYesterdayStartJobNameList;
  initVar.nasTodayStartJobNameList = nasTodayStartJobNameList;
  initVar.nasJobNameList = initVar.nasTodayStartJobNameList.concat(initVar.nasYesterdayStartJobNameList);
  return initVar;
}