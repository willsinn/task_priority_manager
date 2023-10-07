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
            targetCell.setValue(newId);
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
      prioritySheet.getRange(getPrioSheetRowIdx + 1, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
      handleSheetNameChange();
    }
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
        }
    }
