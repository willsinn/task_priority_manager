const main_sheet = "task_priority_manager";
const completed_sheet = "completed_tasks"
const priorityArray = ["Critical", "High", "Medium", "Low"];
const tSheet = (name) => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);


const prioritySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(main_sheet);
const prioritySheetColA = prioritySheet.getRange('A:A').getValues();
const prioritySheetTaskCountCell = prioritySheet.getRange('A2');

const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const activeSheetName = activeSheet.getName();
const activeSheetCellA1Value = activeSheet.getRange('A1').getValue();
const activeCell = activeSheet.getActiveCell();
const activeCellValue = activeCell.getValue();
const activeRowIdx = activeCell.getRowIndex();
const activeColIdx = activeCell.getColumnIndex();
const activeSheetColA_values = activeSheet.getRange('A:A').getValues();
const activeMaxCols = activeSheet.getMaxColumns();

const handleUpdateIDsOnSheetNameChange = () => activeSheetColA_values.map((tId, idx) => { //map through all existing IDs and update them to match the new sheet name
      if (idx > 1 && tId[0]) { //skip first two rows
          const getIdParts = tId[0].split("__", 2);
            const newId = `${activeSheetName}` + '__' + `${getIdParts[1]}`;
            const targetCell = activeSheet.getRange(`A${parseInt(idx+1)}`)
            targetCell.setValue(newId);
      }  
    })

function handleUpdateCellA1() {
      const getSheetNameCell = activeSheet.getRange('A1');
      getSheetNameCell.setValue(activeSheetName);
}

function handleUpdateMainTaskCountCellA2(count) { 
      const newCount = parseInt(count) + 1;
      prioritySheet.getRange('A2').setValue(newCount); //update task counter
}

const handleUpdatePriorityIds = () => prioritySheetColA.map((pId, idx) => {
    if (idx > 1 && pId[0]) {
    const getIdParts = pId[0].split("__", 2);
        if (getIdParts[0] == activeSheetCellA1Value) {
            const newId = `${activeSheetName}` + '__' + `${getIdParts[1]}`;
            const targetCell = prioritySheet.getRange(`A${parseInt(idx+1)}`)
            const projectCell = prioritySheet.getRange(`B${parseInt(idx+1)}`)
            targetCell.setValue(newId);
            projectCell.setValue(activeSheetName)
        }
    }
})

function handleSheetNameChange() {
      if (activeSheetName !== main_sheet && activeSheetName !== activeSheetCellA1Value) {
        handleUpdateIDsOnSheetNameChange();
        handleUpdatePriorityIds();
        handleUpdateCellA1();
      }
    }

function handleCopyNewTaskToPriorityManager(tId) {
    if (tId) {
      const newTaskRow = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols);
      const newTaskValues = newTaskRow.getValues()
      const getPrioSheetRowIdx = prioritySheet.getLastRow();
      prioritySheet.insertRowAfter(getPrioSheetRowIdx); //insert new row
      newTaskValues[0].splice(1, 0, activeSheetName);
      prioritySheet.getRange(getPrioSheetRowIdx + 1, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
    }
}

function handleSyncCellValueByTaskId(newVal) { 
      if (newVal) {
          const valTaskId = activeSheet.getRange('A:A').getValues()[activeSheet.getActiveCell().getRowIndex() - 1][0];
          const valColHeader = activeSheet.getRange(2, activeSheet.getActiveCell().getColumnIndex()).getValue();
          const valProjectName = valTaskId.slice(0, -9); //slice off __#######
          let targetSheet;

          if (activeSheetName === main_sheet) { // IF ANY values on the priority page are edited, this script will UPDATE corresponding PROJECT SHEET values to match the changes
            targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(valProjectName);
          } else { // IF ANY values on the project page are edited, this script will UPDATE the PRIORITY SHEET values to match the changes
            targetSheet = prioritySheet;
          }

          const targetValueTaskIdRow = targetSheet.getRange('A:A').getValues()
          const rowArray = []
          targetValueTaskIdRow.forEach((r, i) => {
            if (r[0] && i > 1 && valTaskId === r[0]) {
              rowArray.push(i);
            }
          });
          const targetSheetMaxCols = targetSheet.getMaxColumns();
          const targetHeaders = targetSheet.getRange(2, 1, 1, targetSheetMaxCols).getValues();
          const targetTaskColIdx = targetHeaders[0].indexOf(valColHeader);
          const targetSheetCell = targetSheet.getRange(rowArray[0]+1, targetTaskColIdx + 1)
          targetSheetCell.setValue(newVal)
        }
}

function handleCreateUniqueTaskID(cell) {
    const taskCount = prioritySheetTaskCountCell.getValue();
        const taskId = `${activeSheetName}` + "__" + `${taskCount}`;
        handleUpdateMainTaskCountCellA2(taskCount)
        cell.setValue(`${taskId}`); //add id to first column
        handleUpdatePriorityIds();
        handleCopyNewTaskToPriorityManager(taskId); 
  }
function handlePriorityLevelChange(prioLvl) {
      const activeIdCell = activeSheet.getRange(`A${activeRowIdx}`);
      const activeIdValue = activeIdCell.getValue();
      if (!activeIdValue) {
          handleCreateUniqueTaskID(activeIdCell)
      } else {
          handleSyncCellValueByTaskId(prioLvl)
      }
}
function alertMessageOKButton() {
  const result = SpreadsheetApp.getUi().alert("Alert message", SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast(result);
} 

function handleCompletedTask() {
  // alertMessageOKButton();
      const completedSheet = tSheet(completed_sheet);
      const completedTaskRow = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols);
      const completedTaskValues = completedTaskRow.getValues()
      const getCompletedSheetLastRowIdx = completedSheet.getLastRow();
      completedSheet.insertRowAfter(getCompletedSheetLastRowIdx); //insert new row
      completedSheet.getRange(getCompletedSheetLastRowIdx + 1, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)
}
function onEdit(e) {
      if (e && activeSheetName === completed_sheet) { // checks if user is editing the completed sheet
      } 
      else {
          handleSheetNameChange();
          if (priorityArray.includes(activeCellValue)) {
            handlePriorityLevelChange(activeCellValue)
          } 
          else if (activeCellValue === true) {
            handleCompletedTask()
          } 
          else {
            handleSyncCellValueByTaskId(activeCellValue)
          }
        }       
  }
    
