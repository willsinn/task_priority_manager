const MAIN_SHEET_NAME = "task_priority_manager";
const COMPLETED_SHEET_NAME = "completed_tasks";
const PRIO_LVLS_COL_NAME = "Priority Level";
const DUE_DATE_COL_NAME = "Due Date";
const COMPLETE_COL_NAME = "Complete"
const INSERT_TASK_ROW_IDX = 3;
const PRIORITY_LEVELS = {
                          'Critical': 0,
                          'High': 1,
                          'Medium': 2,
                          'Low': 3,
                        }
const priorityLevelKeys = Object.keys(PRIORITY_LEVELS)
const activeSS = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const allProjectNames = allSheetNames();
const getSheetBySheetName = (name) => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
const grabActiveCellByLetterNumberValue = (letternumber) => activeSS.getRange(letternumber);
const getCurrentDateTimestampValue = () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME).getRange('I1').getValue().toString();
const getNewActiveCell = () => activeSS.getActiveCell();
const getActiveSheetName = () => activeSS.getName();
const getActiveCellColIdx = () => activeSS.getActiveCell().getColumnIndex();
const getActiveCellRowIdx = () => activeSS.getActiveCell().getColumnIndex();
const getTaskIdByRowIdx = (rowIdx) => activeSS.getRange(`A${ rowIdx }`).getValue();
const getActiveSheetMaxCols = () => activeSS.getMaxColumns();
const getAllRowValuesByMaxCols = (rowIdx, numCols) => activeSS.getRange(rowIdx, 1, 1, numCols).getValues();

const getActiveColHeaderCellByColIdx = (colIdx) => activeSS.getRange(2, colIdx);

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








function handleMoveCompletedRowTaskValuesAcrossSheets(task_id, is_complete, from_sheets_arr, to_sheets_arr, row_arr, invalid_err_msg) {
        let isValidTask = task_id;


        if (!isValidTask) { // conditional to check if the task is real
            const revertCell = getNewActiveCell();
            revertCell.setValue(!is_complete)
            alertMessageOnly(invalid_err_msg)
            return;
        }
        // console.log(fromSheets, toSheets, rowValuesArr)
        // copy 
        if (to_sheets_arr) {
          const fromHeaderCellVals = to_sheets_arr[0][0];
          const fromTaskCellVals = to_sheets_arr[0][1];
          const activeSheetMaxCols = fromTaskCellVals.length+1;
          to_sheets_arr.forEach(name => {
            const sheet = getSheetBySheetName(name);
            const toSheetMaxCols = sheet.getMaxColumns();
            const toSheetHeaders = sheet.getRange(2, 1, 1, toSheetMaxCols).getValues();
            sheet.insertRowBefore(INSERT_TASK_ROW_IDX); //insert new row
              if (activeSheetMaxCols === toSheetMaxCols) {
                  
                  
                  
                  fromTaskCellVals.forEach((val, index) => {
                      const fromHeader = fromHeaderCellVals[index];
                      const xAxis_headerMatchColIdx = toSheetHeaders.findIndex(fromHeader);
                      const yAxis_rowIdx = INSERT_TASK_ROW_IDX;
                      console.log("I'M PAUSING HERE, THIS IS UNTESTED", xAxis_headerMatchColIdx, yAxis_rowIdx)
                      const getCellByXY = sheet.getRange(xAxis_headerMatchColIdx, yAxis_rowIdx)
                      getCellByXY.setValue(val);
                  })
              }
              if (activeSheetMaxCols > toSheetMaxCols) {
                
              }
              if (activeSheetMaxCols < toSheetMaxCols) {
                
              }
          })
        }

}
function onCompleteCheckboxUpdate(isComplete) {
  const activeColumnIndex = getActiveCellColIdx();  
  const activeRowIndex = getActiveCellRowIdx();
  const taskId = getTaskIdByRowIdx(activeRowIndex);

  let isHeaderNamedComplete = getActiveColHeaderCellByColIdx(activeColumnIndex).getValue();
  if (isHeaderNamedComplete && isHeaderNamedComplete === COMPLETE_COL_NAME) { // Double check that the column name is "Complete"
        const activeSheetName = getActiveSheetName();
        const taskProjectName = taskId.slice(0, -9); //slice off __#######
        const maxCols = getActiveSheetMaxCols();

        let fromSheets, toSheets, rowArray, invalidErrMsg;
        const fromRowHeaderValues = getAllRowValuesByMaxCols(2, maxCols);
        const fromRowTaskValues = getAllRowValuesByMaxCols(activeRowIndex, maxCols);
        rowArray = [fromRowHeaderValues[0], fromRowTaskValues[0]]

        if (isComplete === false && activeSheetName === COMPLETED_SHEET_NAME) {
                fromSheets = [COMPLETED_SHEET_NAME];
                toSheets = [MAIN_SHEET_NAME, taskProjectName];
                invalidErrMsg = "Error! Not a valid task, invalid tasks cannot be set to Complete. Please try a different row."
              // handleCompleteTask();
          } else {
              // handleRestoreCompletedTask();
                fromSheets = [MAIN_SHEET_NAME, taskProjectName];
                toSheets = [COMPLETED_SHEET_NAME];
                invalidErrMsg = "Unable to execute command, task must exist to be restored to active sheets."

          }

        handleMoveCompletedRowTaskValuesAcrossSheets(taskId, isComplete, fromSheets, toSheets, rowArray, invalidErrMsg);

  }

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
        const projectSheet = getSheetBySheetName(name);

        projectSheet.insertRowBefore(INSERT_TASK_ROW_IDX)
        newTaskValues[0].splice(1, 1); // REMOVE THE EXTRA PROJECT COLUMN
        
        
        
        const taskCount = prioritySheetTaskCountCell.getValue();
        const taskId = `${name}` + "__" + `${taskCount}`;
        newTaskValues[0].splice(0, 1, taskId); // ADD NEW ID VALUE ONTO SHEET

        projectSheet.getRange(INSERT_TASK_ROW_IDX, 1, 1, newTaskValues[0].length).setValues(newTaskValues)
        const mainTaskIdCell = grabActiveCellByLetterNumberValue(`A${activeRowIdx}`)
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

function handlePriorityLevelChange(prioLvl) {

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
function alertMessageWithOKButton(text) {
  const result = SpreadsheetApp.getUi().alert(`${text}`, SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast(result);
} 
function alertMessageOnly(text) {
  const message = SpreadsheetApp.getUi().alert(`${text}`);
  SpreadsheetApp.getActive().toast(message);
} 
function copyTaskToCompletedSheet(sName) {
      const completedSheet = getSheetBySheetName(COMPLETED_SHEET_NAME);
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
      const sheetIds = getSheetBySheetName(name).getRange('A:A').getValues();
      const errorMsg = "task_id doesn't match on" +`${name}` + "sheet, please report to dev";
      const task = {};
      sheetIds.forEach((tId, index) => {
          if (tId[0] && index > 1 && tId[0] === taskId) {
            task.id = tId[0];
            task.rowIdx = index + 1;
          }
      })
      if (task.id === taskId) {
        getSheetBySheetName(name).deleteRow(task.rowIdx)
      } 
      else {
        alertMessageWithOKButton(errorMsg)
      }
}
function handleCompleteTask() {
      if (activeSheetName !== COMPLETED_SHEET_NAME) {
        const taskId = activeSheet.getRange(`A${ activeRowIdx }`).getValue();
        const projectName = taskId.slice(0, -9); //slice off __#######
        copyTaskToCompletedSheet(projectName);
        deleteCompletedTaskFromSheet(projectName, taskId);
        deleteCompletedTaskFromSheet(MAIN_SHEET_NAME, taskId);
      }
}

function handleRestoreCompletedTask() {
      const completedTaskValues = activeSheet.getRange(activeRowIdx, 1, 1, activeMaxCols).getValues();
      
      if (completedTaskValues[0][0]) {
        prioritySheet.insertRowBefore(INSERT_TASK_ROW_IDX);

        prioritySheet.getRange(INSERT_TASK_ROW_IDX, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)

        completedTaskValues[0].splice(1, 1) //remove the project column
        const projectName = completedTaskValues[0][0].slice(0, -9)
        const projectSheet = getSheetBySheetName(projectName);
        
        projectSheet.insertRowBefore(INSERT_TASK_ROW_IDX); //insert new row
        projectSheet.getRange(INSERT_TASK_ROW_IDX, 1, 1, completedTaskValues[0].length).setValues(completedTaskValues)
        
        getSheetBySheetName(COMPLETED_SHEET_NAME).deleteRow(activeRowIdx) // remove row from completed
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
function onEdit() {
      const newValue = getNewActiveCell().getValue();
      checkForChangesToSheetNames();
console.log(typeof newValue)
      // setLastUpdatedValue(activeSheet, activeMaxCols, activeRowIdx) // updates the ACTIVE SHEET's last updated value
      if (typeof newValue === "boolean") { // checks if user is editing the completed sheet
          onCompleteCheckboxUpdate(newValue); 
      } 
      else {

          if (priorityLevelKeys.includes(newValue)) {
            handlePriorityLevelChange(newValue);
          
          } 
          else {
            handleSyncCellValueByTaskId(newValue);
          
          }
          sortByPriortyThenDueDate(newValue);

        }

  }
    
