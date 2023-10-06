const priorityArray = ["Critical", "High", "Medium", "Low"];


const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const sheetName = activeSheet.getName();
const cellSheetNameValue = activeSheet.getRange('A1').getValue();
const cellTaskCount = activeSheet.getRange('A2');

const activeCell = activeSheet.getActiveCell();
const rowIdx = activeCell.getRowIndex();
    const colA = activeSheet.getRange('A:A').getValues();

const handleUpdateIDsOnSheetNameChange = () => colA.map((tId, idx) => { //map through all existing IDs and update them to match the new sheet name
      const task = { idvalue: "" };

      if (idx > 1 && tId[0]) { //skip first two rows
          let tIdSplit = tId[0].split("");

          for (let i = 0; i < 5; i++) { // account for a 6 digit ID number
              const lastLetter = tIdSplit.pop();
              if (lastLetter == "_" || !parseInt(lastLetter)) { // if last letter is NaN
                  break;
              } else if (!task.idvalue) {
                  task.idvalue = `${lastLetter}`;
              } else {
                  const str = `${lastLetter}${task.idvalue}`
                  task.idvalue = str;
                }
            }

            const newId = `${sheetName}` + '_' + `${task.idvalue}`;
            const targetCell = activeSheet.getRange(`A${parseInt(idx+1)}`)
            targetCell.setValue(newId);
        }
      })

function handleUpdateCellA1() {
      const getSheetNameCell = activeSheet.getRange('A1');
      getSheetNameCell.setValue(sheetName);
      return;
}

function handleSheetNameChange() {
      if (sheetName !== cellSheetNameValue) {
        handleUpdateIDsOnSheetNameChange()
        handleUpdateCellA1();
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
          handleCreateUniqueTaskID()
        } else {
          handleSheetNameChange()
        }
    }
