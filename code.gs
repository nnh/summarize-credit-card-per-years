/**
 * Get the sheet for output.
 * @param none.
 * @return {Object} The sheet object.
 */
function getTargetSpreadSheet_(){
  const url = PropertiesService.getScriptProperties().getProperty('outputSheetUrl');
  return SpreadsheetApp.openByUrl(url);
}
/**
 * Aggregate the total amount for each year.
 * @param none.
 * @return none.
 */
function getAnnualAggregate(){
  // Output on leftmost sheet.
  const outputSheet = getTargetSpreadSheet_().getSheets()[PropertiesService.getScriptProperties().getProperty('datasourceSheetIndex')];
  outputSheet.clearContents();
  // Get input files.
  const targetDriveId = PropertiesService.getScriptProperties().getProperty('targetDriveId');
  const targetDrive = DriveApp.getFolderById(targetDriveId);
  const targetFiles = targetDrive.getFilesByType(MimeType.GOOGLE_SHEETS);
  let outputValues1 = [];
  while (targetFiles.hasNext()){
    const targetFile = targetFiles.next();
    const res = getCreditCardStatement_(targetFile);
    outputValues1 = outputValues1.concat(res);
  }
  outputValues1 = outputValues1.filter(x => x);
  // The items are listed in descending order of total value, and those ranked 11th and below are aggregated as 'Other'.
  const summarizeItemList = getSummarizeItemList_(outputValues1);
  const idxList = getCreditCardValueTableIdx_();
  let outputValues2 = outputValues1.map(outputValue => {
    const temp1 = summarizeItemList.indexOf(outputValue[idxList.item]) > -1 ? 'その他' : outputValue[idxList.item];
    const res1 = outputValue.concat(temp1);
    const temp2 = outputValue[idxList.item] === PropertiesService.getScriptProperties().getProperty('awsName') ? 'AWS' : 'AWS以外';
    const res2 = res1.concat(temp2);
    return res2;
  });
  // Header
  outputValues2.unshift(['年度', '項目', '金額', 'グラフ用項目', 'グラフ用項目（AWS）']);
  outputSheet.getRange(1, 1, outputValues2.length, outputValues2[0].length).setValues(outputValues2);
}
/**
 * Obtain information on the year, item names and total amounts from the 'List' sheet of the 'Credit card statement total' spreadsheet.
 * @param {Object} The file object.
 * @return {Array} Two-dimensional array of [year, item name, total amount].
 */
function getCreditCardStatement_(targetFile){
  const targetSheetName = 'List';
  const targetSheet = SpreadsheetApp.openById(targetFile.getId()).getSheetByName(targetSheetName);
  if (!targetSheet){
    return;
  }
  const targetValues = targetSheet.getDataRange().getValues();
  const colKeiIdx = targetValues[0].indexOf('計');
  const rowKeiIdx = targetValues.map((x, idx) => x[0] === '計' ? idx : null).filter(x => x)[0];
  if (colKeiIdx < 0 | rowKeiIdx < 0){
    return;
  }
  const tempValues = targetValues.filter((_, idx) => 0 <idx & idx < rowKeiIdx);
  // Get the year from the file name.
  const targetFileName = targetFile.getName().substring(0, 4);
  const res = tempValues.map(x => [targetFileName, replaceItemName_(x[0]), x[colKeiIdx]]);
  return res;
}
/**
 * Match item names.
 * @param {String} item name.
 * @return {String} Converted item name.
 */
function replaceItemName_(inputItemName){
  const res = inputItemName === 'AWS.AMAZON.CO' ? PropertiesService.getScriptProperties().getProperty('awsName') :
              inputItemName === 'DOCKER INC. DOCKERIN (WWW.DOCKER.CO)' ? 'DOCKER' :
              inputItemName === 'GITHUB.COM' ? 'GITHUB' :
              inputItemName === 'PIVOTAL SOFTWARE INC.' | inputItemName === 'PIVOTAL' ? 'PIVOTAL TRACKER' :
              inputItemName;
  return res;
}
class CreatePivotTable{
  constructor(){
    const ss = getTargetSpreadSheet_();
    const inputSheet = ss.getSheets()[PropertiesService.getScriptProperties().getProperty('datasourceSheetIndex')];
    this.dataSource = inputSheet.getDataRange();
  }
  /**
   * Create a pivot table in a cell on the specified sheet.
   * @param {Object} Output sheet.
   * @param {Number} Output Row number.
   * @param {Number} Output Column number.
   * @return {Object} Pivot table object created.
   */
  createPivotTable(outputSheet, createRow, createCol){
    const pivotTable = outputSheet.getRange(createRow, createCol).createPivotTable(this.dataSource);
    return pivotTable;
  }
  /**
   * Set up the aggregation method for pivot tables.
   * @param {Object} The pivot table object.
   * @param {Number} Specify the type of aggregation.
   * @param {String} String to be displayed in the leftmost cell.
   * @param {Number} Index of the column to be aggregated.
   * @return {Object} The pivot table object.
   */
  editYearAndAmountTable(pivotTable, pivotType, displayName, colGroupIdx){
    const itemList = getCreditCardValueTableIdx_();
    const rowGroup = pivotTable.addRowGroup(itemList.year + 1);
    rowGroup.showTotals(true).sortAscending;
    const colGroup = pivotTable.addColumnGroup(colGroupIdx + 1);
    const pivotValue = pivotTable.addPivotValue(itemList.amount + 1, pivotType);
    pivotValue.setDisplayName(displayName);
    colGroup.sortDescending().sortBy(pivotTable.getPivotValues()[0], [],);
    return pivotValue;
  }
  /**
   * Create a pivot table for 'totals'.
   * @param {Object} The pivot table object.
   * @param {Number} Specify the type of aggregation.
   * @param {String} String to be displayed in the leftmost cell.
   * @param {Number} Index of the column to be aggregated.
   * @return {Number} Number of the last row of the pivot table output.
   */
  createPivotTableSum(outputSheet, createRow, createCol, colGroupIdx){
    const pivotTable = this.createPivotTable(outputSheet, createRow, createCol);
    const _temp = this.editYearAndAmountTable(pivotTable, SpreadsheetApp.PivotTableSummarizeFunction.SUM, '合計', colGroupIdx);
    const lastRow = outputSheet.getLastRow();
    return lastRow;
  }
  /**
   * Create a pivot table for 'percentage'.
   * @param {Object} The pivot table object.
   * @param {Number} Specify the type of aggregation.
   * @param {String} String to be displayed in the leftmost cell.
   * @param {Number} Index of the column to be aggregated.
   * @return {Number} Number of the last row of the pivot table output.
   */
  createPivotTablePer(outputSheet, createRow, createCol, colGroupIdx){
    const pivotTable = this.createPivotTable(outputSheet, createRow, createCol);
    const pivotValue = this.editYearAndAmountTable(pivotTable, SpreadsheetApp.PivotTableSummarizeFunction.SUM, 'パーセンテージ', colGroupIdx);
    pivotValue.showAs(SpreadsheetApp.PivotValueDisplayType.PERCENT_OF_ROW_TOTAL);
    const lastRow = outputSheet.getLastRow();
    return lastRow;
  }
}
/**
 * Create pivot tables.
 * @param none.
 * @return none.
 */
function createPivotTable(){
  const ss = getTargetSpreadSheet_();
  const pt = new CreatePivotTable();
  const createPivotTableSheetIndex = 1;
  const outputSheet = ss.getSheets()[createPivotTableSheetIndex];
  outputSheet.clearContents();
  const outputStartRow = 1;
  const outputStartCol = 1;
  const interval = 2;
  const itemList = getCreditCardValueTableIdx_();
  const pivotTable1LastRow = pt.createPivotTableSum(outputSheet, outputStartRow, outputStartCol, itemList.itemForSummary);
  const pivot2OutputStartRow = pivotTable1LastRow + interval;
  const pivotTable2LastRow = pt.createPivotTablePer(outputSheet, pivot2OutputStartRow, outputStartCol, itemList.itemForSummary);
  const pivot3OutputStartRow = pivotTable2LastRow + interval;
  const pivotTable3LastRow = pt.createPivotTableSum(outputSheet, pivot3OutputStartRow, outputStartCol, itemList.itemForSummaryAws);
  const pivot4OutputStartRow = pivotTable3LastRow + interval;
  const _temp = pt.createPivotTablePer(outputSheet, pivot4OutputStartRow, outputStartCol, itemList.itemForSummaryAws);
}
function getCreditCardValueTableIdx_(){
  let creditCardValueIdx ={};
  creditCardValueIdx.year = 0;
  creditCardValueIdx.item = 1;
  creditCardValueIdx.amount = 2;
  creditCardValueIdx.itemForSummary = 3;
  creditCardValueIdx.itemForSummaryAws = 4;
  return creditCardValueIdx;
}
/**
 * Summarising 11th place and below.
 * @param {Array} Two-dimensional array of [year, item name, total amount].
 * @return {Array} Array of item names to be aggregated.
 */
function getSummarizeItemList_(targetValues){
  const summarizeStartIdx = 11;
  const idxList = getCreditCardValueTableIdx_();
  const itemList = targetValues.map((x, idx) => idx > 0 ? x[idxList.item] :null).filter(x => x);
  const uniqueItemList = Array.from(new Set(itemList));
  const targetValuesGroupByItem = uniqueItemList.map(item => targetValues.filter(x => x[idxList.item] === item));
  if (targetValuesGroupByItem.length < summarizeStartIdx) {
    return null;
  }
  const tempTotalList = targetValuesGroupByItem.map(x => {
    const totalAmount = x.reduce((total, current) => total + current[idxList.amount], 0); 
    return [x[0][idxList.item], totalAmount];
  });
  const itemIdx = 0;
  const amountIdx = 1;
  const descTotalList = tempTotalList.sort((a, b) => b[amountIdx] - a[amountIdx]);
  const summarizeTargetItems = descTotalList.filter((_, idx) => idx >= summarizeStartIdx).map(x => x[itemIdx]);
  return summarizeTargetItems;
}