const mainSheet = "task_priority_manager";
const priorityArray = ["Critical", "High", "Medium", "Low"];

const prioritySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheet);
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
      if (activeSheetName !== mainSheet && activeSheetName !== activeSheetCellA1Value) {
        handleUpdateIDsOnSheetNameChange();
        handleUpdatePriorityIds();
        handleUpdateCellA1();
      }
    }

function handleCopyNewTaskToPriorityManager(tId) {
    if (tId) {
      const numCols = activeSheet.getMaxColumns();
      const newTaskRow = activeSheet.getRange(activeRowIdx, 1, 1, numCols);
      const newTaskValues = newTaskRow.getValues()
      const getPrioSheetRowIdx = prioritySheet.getLastRow();
      prioritySheet.insertRowAfter(getPrioSheetRowIdx); //insert new row
      newTaskValues[0].splice(1, 0, activeSheetName);
      prioritySheet.getRange(getPrioSheetRowIdx + 1, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
    }
}

function updatePriorityTaskValue(val) { // IF ANY values on the project page are edited, this script will UPDATE the PRIORITY SHEET values to match the changes

}

function updateProjectTaskValue(val) { // IF ANY values on the priority page are edited, this script will UPDATE corresponding PROJECT SHEET values to match the changes
        const valueTaskId = activeSheetColA_values[activeRowIdx - 1][0];
        const valueColHeader = activeSheet.getRange(2, activeColIdx).getValue();
        const valueProjectName = valueTaskId.slice(0, -9); //slice off __#######
        const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(valueProjectName);
        const targetValueTaskIdRow = targetSheet.getRange('A:A').getValues()
        const rowArray = []
        targetValueTaskIdRow.forEach((r, i) => {
          if (r[0] && i > 1 && valueTaskId === r[0]) {
            rowArray.push(i);
          }
        });
        const targetSheetMaxCols = targetSheet.getMaxColumns();
        const targetHeaders = targetSheet.getRange(2, 1, 1, targetSheetMaxCols).getValues();
        const taskColIdx = targetHeaders[0].indexOf(valueColHeader);
        const tCell = targetSheet.getRange(rowArray[0]+1, taskColIdx + 1)
        tCell.setValue(val)
}
function handleSyncCellValueByTaskId(newVal) {
    if (newVal) {
        if (activeSheetName === mainSheet) {
          updateProjectTaskValue(newVal);
        } else {
          updatePriorityTaskValue(newVal)
        }
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
          Logger.log(activeCellValue)
      } else {
          handleSyncCellValueByTaskId(prioLvl)
      }
}

function onEdit(e) { 
      handleSheetNameChange();
      if (e && priorityArray.includes(activeCellValue)) {
          handlePriorityLevelChange(activeCellValue)
        } else {
          handleSyncCellValueByTaskId(activeCellValue)
        }
    }
