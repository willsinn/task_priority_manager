const main_sheet = "task_priority_manager";
const completed_sheet = "completed_tasks";
const priorityArray = ["Critical", "High", "Medium", "Low"];
const insertRowIdx = 3;
const projectNames = allSheetNames();



const targSheet = (name) => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
const grabActiveCell = (str) => SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(str);
const currDateTime = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(main_sheet).getRange('I1').getValue().toString();


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




function allSheetNames() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  const namesArr = [];
  sheets.forEach((sheet) => {
    const name = sheet.getName();
    if (main_sheet !== name && completed_sheet !== name) {
      namesArr.push(name);
    }
  });
  return namesArr;
}
const handleUpdateIDsOnSheetNameChange = () => activeSheetColA_values.map((tId, idx) => { //map through all existing IDs and update them to match the new sheet name
      if (idx > 1 && tId[0]) { //skip first two rows
          const getIdParts = tId[0].split("__", 2);
            const newId = `${activeSheetName}` + '__' + `${ getIdParts[1] }`;
            const targetCell = activeSheet.getRange(`A${ parseInt( idx + 1 ) }`)
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
            const newId = `${activeSheetName}` + '__' + `${ getIdParts[1] }`;
            const targetCell = prioritySheet.getRange(`A${ parseInt( idx+1 ) }`)
            const projectCell = prioritySheet.getRange(`B${ parseInt( idx+1 ) }`)
            targetCell.setValue(newId);
            projectCell.setValue(activeSheetName)
        }
    }
})

function handleSheetNameChange() {
      if (activeSheetName !== completed_sheet && activeSheetName !== main_sheet && activeSheetName !== activeSheetCellA1Value) {
        handleUpdateIDsOnSheetNameChange();
        handleUpdatePriorityIds();
        handleUpdateCellA1();
      }
    }

function handleCopyNewTaskToPriorityManager(tId) {
    if (tId) {
      const newTaskRow = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols) ;
      const newTaskValues = newTaskRow.getValues()
      prioritySheet.insertRowBefore(insertRowIdx); //insert new row
      newTaskValues[0].splice(1, 0, activeSheetName);
      prioritySheet.getRange(insertRowIdx, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
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
          setLastUpdatedValue(targetSheet, targetSheetMaxCols, rowArray[0]+1) // updates the TARGET SHEET's last updated value
        }
}

function copyNewTaskToProjectSheet(name) {
        const newTaskValues = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols).getValues();
        const projectSheet = targSheet(name);

        projectSheet.insertRowBefore(insertRowIdx)
        newTaskValues[0].splice(1, 1); // REMOVE THE EXTRA PROJECT COLUMN
        
        
        
        const taskCount = prioritySheetTaskCountCell.getValue();
        const taskId = `${name}` + "__" + `${taskCount}`;
        newTaskValues[0].splice(0, 1, taskId); // ADD NEW ID VALUE ONTO SHEET

        projectSheet.getRange(insertRowIdx, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
        const mainTaskIdCell = grabActiveCell(`A${activeRowIdx}`)
        mainTaskIdCell.setValue(taskId)

        handleUpdateMainTaskCountCellA2(taskCount)


}

function handleMainSheetCreateTask() {
      const errorMsg = `Error! Could not create task due to invalid project name in cell B-${activeRowIdx}. The value needs to be an exact match to ONE sheet's name, please update cell B-${activeRowIdx} with one of following sheet names to continue: ` + `'${projectNames.join("',  '")}'`
      const currCellVal = activeSheet.getRange(`B${activeRowIdx}`).getValue();

      if (!projectNames.includes(currCellVal)) {
        activeCell.setValue("")
        alertMessageWithOKButton(errorMsg)
      } else {
        copyNewTaskToProjectSheet(currCellVal);
      }
}
function createUniqueTaskId(cell) {
    const taskCount = prioritySheetTaskCountCell.getValue();
        const taskId = `${activeSheetName}` + "__" + `${taskCount}`;
        handleUpdateMainTaskCountCellA2(taskCount)
        cell.setValue(`${taskId}`); //add id to first column
        handleUpdatePriorityIds();
        handleCopyNewTaskToPriorityManager(taskId); 
  }

function setLastUpdatedValue(sheet, maxCols, rowIdx) {
    if (rowIdx > 2) {
      const lastUpdatedCol = sheet.getRange(rowIdx, maxCols)
      lastUpdatedCol.setValue(currDateTime)
    }
}

function handlePriorityLevelChange(prioLvl) {
      const activeIdCell = activeSheet.getRange(`A${activeRowIdx}`);
      const activeIdValue = activeIdCell.getValue();
      if (!activeIdValue) {
          if (activeSheetName === main_sheet) {
            handleMainSheetCreateTask();
          } else {
            createUniqueTaskId(activeIdCell)
          }
      } else {
          handleSyncCellValueByTaskId(prioLvl)
      }
}
function alertMessageWithOKButton(val) {
  const result = SpreadsheetApp.getUi().alert(`${val}`, SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast(result);
} 
function copyTaskToCompletedSheet(sName) {
      const completedSheet = targSheet(completed_sheet);
      const completedTaskValues = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols).getValues();
      completedSheet.insertRowBefore(insertRowIdx); //insert new row
            
      if (sName === activeSheetName) {
        completedTaskValues[0].splice(1, 0, activeSheetName);
      } 

      completedSheet.getRange(insertRowIdx, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)
}
function deleteCompletedTaskFromSheet(name, taskId) {
      const sheetIds = targSheet(name).getRange('A:A').getValues();
      const errorMsg = "task_id doesn't match on" +`${name}` + "sheet, please report to dev";
      const task = {};
      sheetIds.forEach((tId, index) => {
          if (tId[0] && index > 1 && tId[0] === taskId) {
            task.id = tId[0];
            task.rowIdx = index + 1;
          }
      })
      if (task.id === taskId) {
        targSheet(name).deleteRow(task.rowIdx)
      } 
      else {
        alertMessageWithOKButton(errorMsg)
      }
}
function handleCompleteTask() {
      if (activeSheetName !== completed_sheet) {
        const taskId = activeSheet.getRange(`A${ activeRowIdx }`).getValue();
        const projectName = taskId.slice(0, -9); //slice off __#######
        copyTaskToCompletedSheet(projectName);
        deleteCompletedTaskFromSheet(projectName, taskId);
        deleteCompletedTaskFromSheet(main_sheet, taskId);
      }
}

function handleRestoreCompletedTask() {
      const completedTaskValues = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols).getValues();
      
      if (completedTaskValues[0][0]) {
        prioritySheet.insertRowBefore(insertRowIdx);

        prioritySheet.getRange(insertRowIdx, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)

        completedTaskValues[0].splice(1, 1) //remove the project column
        const projectName = completedTaskValues[0][0].slice(0, -9)
        const projectSheet = targSheet(projectName);
        
        projectSheet.insertRowBefore(insertRowIdx); //insert new row
        projectSheet.getRange(insertRowIdx, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)
        
        targSheet(completed_sheet).deleteRow(activeRowIdx) // remove row from completed
      }
}
function sortByPriortyThenDueDate() {
      const headerValue = activeSheet.getRange(2, activeSheet.getActiveCell().getColumnIndex()).getValue();

      if (priorityArray.includes(activeCellValue) || headerValue === "Due Date") {
        const priority = {
            critical: 0,
            high: 1,
            medium: 2,
            low: 3,
          }
        const data = prioritySheet.getRange('A3:I').getValues();
        data.sort((a, b) => { // SORTS BY PRIORITY LEVEL THEN DATE.
          const aOrder = priority[a[3].toLowerCase()];
          const bOrder = priority[b[3].toLowerCase()];
          const aDueDate = a[6];
          const bDueDate = b[6];
              let n = aOrder - bOrder;
              if (n !== 0) {
                  return n;
              }

              /* // COMMENT THIS CONDITIONAL IN TO GROUP ALL DATELESS TASKS BELOW DATED TASKS, KEEP IT COMMENTED OUT TO GROUP DATELESS TASKS ABOVE DATED TASKS 
                if(!bDueDate || !aDueDate){
                return bDueDate - aDueDate
              } 
              */

              return aDueDate - bDueDate;
        })
        prioritySheet.getRange('A3:I').setValues(data);
      }
}
function onEdit(e) {
      setLastUpdatedValue(activeSheet, activeMaxCols, activeRowIdx) // updates the ACTIVE SHEET's last updated value
      if (e && activeSheetName === completed_sheet && activeCellValue === false) { // checks if user is editing the completed sheet
          handleRestoreCompletedTask(); 
      } 
      else {
          handleSheetNameChange();
          if (priorityArray.includes(activeCellValue)) {
            handlePriorityLevelChange(activeCellValue);
          
          } 
          else if (activeCellValue === true) {
            handleCompleteTask();
          
          } 
          else {
            handleSyncCellValueByTaskId(activeCellValue);
          
          }
          sortByPriortyThenDueDate();

        }

  }
    
