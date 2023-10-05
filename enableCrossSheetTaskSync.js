const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const sheetName = activeSheet.getName();
const cellSheetName = activeSheet.getRange('A1');
const cellTaskCount = activeSheet.getRange('A2');

const activeCell = activeSheet.getActiveCell();
    const rowIdx = activeCell.getRowIndex();



function setSheetName(e) { // 1B. Set up logic to return the name of the sheet as a string value.
    return activeSheet.getRange('A1').setValue(sheetName);
  }

function handleUpdateActiveSheetTaskIds() { // 1B+1. FOR IF SHEET NAME IS CHANGED, create logic to update all the ids on both the priority and project page.
    const colA = activeSheet.getRange('A:A').getValues();

    for (let i = 2; i < colA.length; i++) {
        const taskId = colA[i][0];
        Logger.log(taskId[1], "TaskId, Line 20")
        const strLength = taskId.length

        if (!strLength && strLength > 0) {
          const numArray = [];

            for (let j = 0; j < strLength; j++) {
                const letter = taskId.charAt(strLength-1-j);
                Logger.log(letter, "Letter, Line 25")
                if (letter == "_") {
                  break;
                } 
                numArray.push(letter);
            }
          const idNum = numArray.reverse().join("");
          const newId = `${sheetName}` + '_' + `${idNum}`;
          const targetCell = activeSheet.getRange(`A${i}`)
          targetCell.setValue(newId);
          return
        }
    }
  }
  function handleCreateUniqueTaskID() {
    const taskIdCell = activeSheet.getRange(`A${rowIdx}`);

    if (!taskIdCell.getValue()) {
        const taskId = `${cellSheetName.getValue()}` + "_" + `${cellTaskCount.getValue()}`;

        activeSheet.getRange('A2').setValue(cellTaskCount.getValue() +1); //update task counter

        taskIdCell.setValue(`A${taskId}`); //add id to first column
        return;
    } 
  }
function checkSheetName() {

    if (sheetName !== cellSheetName.getValue()) {
      handleUpdateActiveSheetTaskIds()
      setSheetName()
      return;
    }
  }


function onEdit(e) { 
    if (e && activeCell.getValue() === "Critical" || activeCell.getValue() === "High" || activeCell.getValue() === "Medium" || activeCell.getValue() === "Low") {
        checkSheetName();
        handleCreateUniqueTaskID();
        return;
    } 
  }
