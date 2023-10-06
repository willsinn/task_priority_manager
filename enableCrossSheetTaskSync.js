const priorityArray = ["Critical", "High", "Medium", "Low"];


const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const sheetName = activeSheet.getName();
const cellSheetNameValue = activeSheet.getRange('A1').getValue();
const cellTaskCount = activeSheet.getRange('A2');

const activeCell = activeSheet.getActiveCell();
const rowIdx = activeCell.getRowIndex();





function handleSheetNameChange(e) { // 1B+1. FOR IF SHEET NAME IS CHANGED, create logic to update all the ids on both the priority and project page.
    const colA = activeSheet.getRange('A:A').getValues();

    if (colA) { 

    for (let i = 2; i < colA.length; i++) {
        const taskId = colA[i];
        const task = { idvalue: "" };
        if (taskId[0]) {
            
            let isNum = true;
            let idx = taskId[0].length - 1;
            while (isNum) {
                const char = taskId[0][idx];
                if (parseInt(char) === NaN) {
                    isNum = false;
                } else {
                    if (task.idvalue) {
                      const str = `${char}${task.idvalue}`
                      task.idvalue = str;
                    } else {
                      task.idvalue = `${char}`;
                  }
                }
                idx--;
              }
            const newId = `${sheetName}` + '_' + `${task.idvalue}`;
            const targetCell = activeSheet.getRange(`A${parseInt(i+1)}`)
            targetCell.setValue(newId);
        }
    }
    activeSheet.getRange('A1').setValue(sheetName)
  }
}
  
  function handleCreateUniqueTaskID() {
    const taskIdCell = activeSheet.getRange(`A${rowIdx}`);

    if (!taskIdCell.getValue()) {
        const taskId = `${cellSheetNameValue}` + "_" + `${cellTaskCount.getValue()}`;

        activeSheet.getRange('A2').setValue(cellTaskCount.getValue() +1); //update task counter

        taskIdCell.setValue(`${taskId}`); //add id to first column
    } 
  }



function onEdit(e) { 
    const activeCellValue = activeCell.getValue();
      if (e && priorityArray.includes(activeCellValue)) {
          handleCreateUniqueTaskID();
        } else {
          handleSheetNameChange()
        }
    }
