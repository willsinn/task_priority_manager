/* GOAL: Duplicate project-sheet task to priority-sheet and prepare task to be synced across sheets.

1. Develop logic to create a unique task_id that can be easily traced back to the project sheet and vice versa.
    1A. Set up a cell that keeps track of the sheet's total task count.
    1B. Set up logic to return the name of the sheet as a string value.
        ** DO LATER ** 1B+1. FOR IF SHEET NAME IS CHANGED, create logic to update all the ids on both the priority and project page.
    1C. Concatenate the task count after the sheet name after an underscore, important rules to include below:
        1C+1. All alphanumeric values must be changed to lowercase before creating the id.
        ** MAY NOT NEED ** 1C+2. All spaces(if any) need to be removed from the sheet name before creating the id.
        1C+2. Value in task count cell needs to be incremented by 1.
    1D. Add the new unique id value to the same row where the onChange event was activated.
2. Create a new row on the priority page and copy the existing row values over.
    2A. Copy the event's row values, 
3. Set up an onChange event on the priority column, the script to create a new task should trigger if the value of the cell changes at all.


*/
