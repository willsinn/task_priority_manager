const MAIN_SHEET_NAME = "task_priority_manager";
const COMPLETED_SHEET_NAME = "completed_tasks";
const PRIO_LVLS_COL_NAME = "Priority Level";
const DUE_DATE_COL_NAME = "Due Date";
const INSERT_TASK_ROW_IDX = 3;
const PRIORITY_LEVELS = {
                          'Critical': 0,
                          'High': 1,
                          'Medium': 2,
                          'Low': 3,
                        }
const priorityLevelKeys = Object.keys(PRIORITY_LEVELS)

const allProjectNames = allSheetNames();
const targSheet = (name) => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
const grabActiveCell = (str) => SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(str);
const getCurrentDateTimestampValue = () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME).getRange('F1').getValue().toString();
const getNewActiveCellValue = () => SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getValue();


const prioritySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
const prioritySheetColA = prioritySheet.getRange('A:A').getValues();
const prioritySheetTaskCountCell = prioritySheet.getRange('A2');

const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const activeSheetName = activeSheet.getName();
const activeSheetCellA1Value = activeSheet.getRange('A1').getValue();
const activeCell = activeSheet.getActiveCell();
const activeRowIdx = activeCell.getRowIndex();
const activeColIdx = activeCell.getColumnIndex();
const activeSheetColA_values = activeSheet.getRange('A:A').getValues();
const activeMaxCols = activeSheet.getMaxColumns();
const activeIdCell = activeSheet.getRange(`A${activeRowIdx}`);
const activeIdValue = activeIdCell.getValue();




function moveRowValuesToAnotherSheet(fromSheet, toSheet, rowValuesArr) {

}
function handleCompleteInputValueActions(value) {
      // let isComplete = ;
      
}
function allSheetNames() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  const namesArr = [];
  sheets.forEach((sheet) => {
    const name = sheet.getName();
    if (MAIN_SHEET_NAME !== name && COMPLETED_SHEET_NAME !== name) {
      namesArr.push(name);
    }
  });
  return namesArr;
}
const updateProjectIdsToNewSheetName = () => activeSheetColA_values.map((tId, idx) => { //map through all existing IDs and update them to match the new sheet name
      if (idx > 1 && tId[0]) { //skip first two rows
          const getIdParts = tId[0].split("__", 2);
            const newId = `${activeSheetName}` + '__' + `${ getIdParts[1] }`;
            const targetCell = activeSheet.getRange(`A${ parseInt( idx + 1 ) }`)
            targetCell.setValue(newId);
      }  
    })

function handleUpdateMainTaskCountCellA2(count) { 
      const newCount = parseInt(count) + 1;
      prioritySheet.getRange('A2').setValue(newCount); //update task counter
}

const updatePrioritySheetIdsToMatchNewSheetName = () => prioritySheetColA.map((pId, idx) => {
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

function checkForChangesToSheetNames() {
      if (activeSheetName !== COMPLETED_SHEET_NAME && activeSheetName !== MAIN_SHEET_NAME && activeSheetName !== activeSheetCellA1Value) {
        updateProjectIdsToNewSheetName();
        updatePrioritySheetIdsToMatchNewSheetName();
        
        const getSheetNameCell = activeSheet.getRange('A1');
        getSheetNameCell.setValue(activeSheetName);
      }
    }

function handleCopyNewTaskToPriorityManager(tId) {
    if (tId) {
      prioritySheet.insertRowBefore(INSERT_TASK_ROW_IDX); //insert new row
      let isCopyInProgress = prioritySheet.getRange(`A${INSERT_TASK_ROW_IDX}`).getValue();
      if (!isCopyInProgress) {
        const newTaskRow = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols) ;
        const newTaskValues = newTaskRow.getValues()
        newTaskValues[0].splice(1, 0, activeSheetName);

        prioritySheet.getRange(INSERT_TASK_ROW_IDX, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
      }
    }
}

function handleSyncCellValueByTaskId(newVal) { 
      if (newVal) {
          const valTaskId = activeSheet.getRange('A:A').getValues()[activeSheet.getActiveCell().getRowIndex() - 1][0];
          const valColHeader = activeSheet.getRange(2, activeSheet.getActiveCell().getColumnIndex()).getValue();
          const valProjectName = valTaskId.slice(0, -9); //slice off __#######
          let targetSheet;

          if (activeSheetName === MAIN_SHEET_NAME) { // IF ANY values on the priority page are edited, this script will UPDATE corresponding PROJECT SHEET values to match the changes
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

        projectSheet.insertRowBefore(INSERT_TASK_ROW_IDX)
        newTaskValues[0].splice(1, 1); // REMOVE THE EXTRA PROJECT COLUMN
        
        
        
        const taskCount = prioritySheetTaskCountCell.getValue();
        const taskId = `${name}` + "__" + `${taskCount}`;
        newTaskValues[0].splice(0, 1, taskId); // ADD NEW ID VALUE ONTO SHEET

        projectSheet.getRange(INSERT_TASK_ROW_IDX, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
        const mainTaskIdCell = grabActiveCell(`A${activeRowIdx}`)
        mainTaskIdCell.setValue(taskId)

        handleUpdateMainTaskCountCellA2(taskCount)

}

function handleMainSheetCreateTask() {
      const errorMsg = `Error! Could not create task due to invalid project name in cell B-${activeRowIdx}. The value needs to be an exact match to ONE sheet's name, please update cell B-${activeRowIdx} with one of following sheet names to continue: ` + `'${allProjectNames.join("',  '")}'`
      const currCellVal = activeSheet.getRange(`B${activeRowIdx}`).getValue();

      if (!allProjectNames.includes(currCellVal)) {
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
        updatePrioritySheetIdsToMatchNewSheetName();
        handleCopyNewTaskToPriorityManager(taskId); 
  }

function setLastUpdatedValue(sheet, maxCols, rowIdx) {
    let timestamp = getCurrentDateTimestampValue();
    if (timestamp && rowIdx > 2) {
      const lastUpdatedCol = sheet.getRange(rowIdx, maxCols)
      lastUpdatedCol.setValue(timestamp)
    }
}

function onChangeTaskPriorityLevel(prioLvl) {

      if (!activeIdValue) {
          if (activeSheetName === MAIN_SHEET_NAME) {
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
      const completedSheet = targSheet(COMPLETED_SHEET_NAME);
      const completedTaskValues = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols).getValues();
      completedSheet.insertRowBefore(INSERT_TASK_ROW_IDX); //insert new row
      let isCopyInProgress = completedSheet.getRange(`A${INSERT_TASK_ROW_IDX}`).getValue();
      if (sName === activeSheetName) {
        completedTaskValues[0].splice(1, 0, activeSheetName);
      } 
      if (!isCopyInProgress) {       

        completedSheet.getRange(INSERT_TASK_ROW_IDX, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)
      }
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
function onSetTaskComplete() {
      if (activeSheetName !== COMPLETED_SHEET_NAME) {
        const taskId = activeSheet.getRange(`A${ activeRowIdx }`).getValue();
        const projectName = taskId.slice(0, -9); //slice off __#######
        copyTaskToCompletedSheet(projectName);
        deleteCompletedTaskFromSheet(projectName, taskId);
        deleteCompletedTaskFromSheet(MAIN_SHEET_NAME, taskId);
      }
}

function onSetTaskIncomplete() {
      const completedTaskValues = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols).getValues();
      
      if (completedTaskValues[0][0]) {
        prioritySheet.insertRowBefore(INSERT_TASK_ROW_IDX);

        prioritySheet.getRange(INSERT_TASK_ROW_IDX, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)

        completedTaskValues[0].splice(1, 1) //remove the project column
        const projectName = completedTaskValues[0][0].slice(0, -9)
        const projectSheet = targSheet(projectName);
        
        projectSheet.insertRowBefore(INSERT_TASK_ROW_IDX); //insert new row
        projectSheet.getRange(INSERT_TASK_ROW_IDX, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)
        
        targSheet(COMPLETED_SHEET_NAME).deleteRow(activeRowIdx) // remove row from completed
      }
}
function sortByPriortyThenDueDate(val) {
      const headerValue = activeSheet.getRange(2, activeSheet.getActiveCell().getColumnIndex()).getValue();

      if (priorityLevelKeys.includes(val) || headerValue === "Due Date") {
        const headers = prioritySheet.getRange(2, 1, 1, activeMaxCols).getValues().flat();
        const prioIdx = headers.indexOf(PRIO_LVLS_COL_NAME);
        const dueDateIdx = headers.indexOf(DUE_DATE_COL_NAME)
        const data = prioritySheet.getRange('A3:I').getValues();
        data.sort((a, b) => { // SORTS BY PRIORITY LEVEL THEN DATE.
          const aOrder = PRIORITY_LEVELS[a[prioIdx]];
          const bOrder = PRIORITY_LEVELS[b[prioIdx]];
          const aDueDate = a[dueDateIdx];
          const bDueDate = b[dueDateIdx];
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
      let newValue = getNewActiveCellValue();
      checkForChangesToSheetNames();

      if (e && activeSheetName === COMPLETED_SHEET_NAME && newValue === false) { // checks if user is editing the completed sheet
          onSetTaskIncomplete(); 
      } 
      else {
          setLastUpdatedValue(activeSheet, activeMaxCols, activeRowIdx) // updates the ACTIVE SHEET's last updated value, this script is put here so it can get deleted by onSetTaskComplete

          if (priorityLevelKeys.includes(newValue)) {
            onChangeTaskPriorityLevel(newValue);
          
          } 
          else if (newValue === true) {
            onSetTaskComplete();
          
          } 
          else {
            handleSyncCellValueByTaskId(newValue);
          
          }
          sortByPriortyThenDueDate(newValue);

        }

  }
    
