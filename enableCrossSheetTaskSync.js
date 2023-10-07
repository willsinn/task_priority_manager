const prioritySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('priority_manager');
const priorityArray = ["Critical", "High", "Medium", "Low"];
const priorityColA = prioritySheet.getRange('A:A').getValues();


const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const sheetName = activeSheet.getName();
const cellSheetNameValue = activeSheet.getRange('A1').getValue();
const cellTaskCount = activeSheet.getRange('A2');

const activeCell = activeSheet.getActiveCell();
const rowIdx = activeCell.getRowIndex();
const colA = activeSheet.getRange('A:A').getValues();

const handleUpdateIDsOnSheetNameChange = () => colA.map((tId, idx) => { //map through all existing IDs and update them to match the new sheet name
      if (idx > 1 && tId[0]) { //skip first two rows
          const getIdParts = tId[0].split("__", 2);
            const newId = `${sheetName}` + '__' + `${getIdParts[1]}`;
            const targetCell = activeSheet.getRange(`A${parseInt(idx+1)}`)
            targetCell.setValue(newId);
      }  
    })


function handleUpdateCellA1() {
      const getSheetNameCell = activeSheet.getRange('A1');
      getSheetNameCell.setValue(sheetName);
}

const handleUpdatePriorityIds = () => priorityColA.map((pId, idx) => {
    if (idx > 1 && pId[0]) {
    const getIdParts = pId[0].split("__", 2);
        if (getIdParts[0] == cellSheetNameValue) {
            const newId = `${sheetName}` + '__' + `${getIdParts[1]}`;
            const targetCell = prioritySheet.getRange(`A${parseInt(idx+1)}`)
            const projectCell = prioritySheet.getRange(`B${parseInt(idx+1)}`)
            targetCell.setValue(newId);
            projectCell.setValue(sheetName)
        }
    }
})

function handleSheetNameChange() {
      if (sheetName !== cellSheetNameValue) {
        handleUpdateIDsOnSheetNameChange();
        handleUpdatePriorityIds();
        handleUpdateCellA1();
      }
    }

function handleCopyNewTaskToPriorityManager(tId) {
    if (tId) {
      const numCols = activeSheet.getMaxColumns();
      const newTaskRow = activeSheet.getRange(rowIdx, 1, 1, numCols);
      const newTaskValues = newTaskRow.getValues()
      const getPrioSheetRowIdx = prioritySheet.getLastRow();
      prioritySheet.insertRowAfter(getPrioSheetRowIdx); //insert new row
      newTaskValues[0].splice(1, 0, sheetName);
      prioritySheet.getRange(getPrioSheetRowIdx + 1, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
      handleSheetNameChange();
    }
}
function syncCellValueByTaskId(newVal) {
    const activeColA = activeSheet.getRange('A:A').getValues();
    const valueTaskId = activeColA[rowIdx-1];
    const valueColHeader = activeSheet.getRange(2, colIdx).getValue();
    const valueProjectName = valueTaskId[0].split("__")[0]
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(valueProjectName);
    const targetValueTaskIdRow = targetSheet.getRange('A:A').getValues()
    const tRowIdx = []
    targetValueTaskIdRow.forEach((r, i) => {
      if (r[0] && i > 1 && valueTaskId[0] === r[0]) {
        tRowIdx.push(i);
        return
      }
    });
    const targetSheetMaxCols = targetSheet.getMaxColumns();
    const targetHeaders = targetSheet.getRange(2, 1, 1, targetSheetMaxCols).getValues();
    const tColIdx = targetHeaders[0].indexOf(valueColHeader);
    const tCell = targetSheet.getRange(tRowIdx[0]+1, tColIdx+1)
    tCell.setValue(newVal)
    Logger.log(tColIdx)
}
function handleCreateUniqueTaskID() {
    const taskIdCell = activeSheet.getRange(`A${rowIdx}`);

    if (!taskIdCell.getValue()) {
        const taskId = `${sheetName}` + "__" + `${cellTaskCount.getValue()}`;
        activeSheet.getRange('A2').setValue(cellTaskCount.getValue() +1); //update task counter
        taskIdCell.setValue(`${taskId}`); //add id to first column
                handleUpdatePriorityIds();
        handleCopyNewTaskToPriorityManager(taskId);
    } else {
        handleSheetNameChange()
    }
  }


function onEdit(e) { 
    const activeCellValue = activeCell.getValue();
      if (e && priorityArray.includes(activeCellValue)) {
          handleCreateUniqueTaskID()
        } else {
          handleSheetNameChange()
            syncCellValueByTaskId(activeCell.getValue())
        }
    }
