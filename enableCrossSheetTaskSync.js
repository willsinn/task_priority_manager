const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const sheetName = activeSheet.getName();
const cellSheetName = activeSheet.getRange('A1');
const cellTaskCount = activeSheet.getRange('A2');

const activeCellValue = activeSheet.getActiveCell().getValue();
function setSheetName(e) { // 1B. Set up logic to return the name of the sheet as a string value.
    return activeSheet.getRange('A1').setValue(sheetName);
  }

  function handleUpdateTaskIds() { // 1B+1. FOR IF SHEET NAME IS CHANGED, create logic to update all the ids on both the priority and project page.
    return console.log("Update Task_IDs");
  }
function checkSheetName() {
    if (sheetName !== cellSheetName.getValue()) {
      setSheetName();
      handleUpdateTaskIds();
      return;
    }
}
function createUniqueTaskID() {
    const taskId = `${cellSheetName.getValue()}` + "_" + `${cellTaskCount.getValue()}`;
    activeSheet.getRange('A2').setValue(cellTaskCount.getValue() +1)
    return taskId;
}
function onEdit(e) {  
  if (e && activeCellValue === 'Critical' || activeCellValue === 'High') {
        // NEED TO CHECK THAT AN ID DOESN'T ALREADY EXIST
        const uniqueTaskId = createUniqueTaskID();
        // NEED TO FIND COLUMN A OF THE SAME ROW AS THE ACTIVE CELL
        return activeSheet.getRange('A10').setValue(uniqueTaskId);
  } else {
  }

}
