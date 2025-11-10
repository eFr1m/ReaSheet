/**
 * This is an example file demonstrating how to use the ReaSheets framework.
 *
 * To run this in a Google Apps Script project, you would need three files:
 * 1. core.js (contains the component definitions)
 * 2. engine.js (contains the render engine)
 * 3. main.js (this file)
 *
 * You would then select the 'createToDoListSheet' function in the Apps Script
 * editor and click "Run".
 */

/**
 * The main function to generate the sheet.
 * This function gets the target sheet, defines the layout using components,
 * and then calls the render function to build the sheet.
 */
function createToDoListSheet() {
    // 1. Get the target sheet object from Google Apps Script
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ToDoList') || ss.insertSheet('ToDoList');
    
    // It's good practice to clear the sheet for a fresh render
    sheet.clear();
    sheet.setFrozenRows(1);

    // 2. Define some reusable styles for clarity
    const headerStyle = Style({
        backgroundColor: '#4a86e8',
        font: { color: 'white', bold: true },
        alignment: { horizontal: 'center', vertical: 'middle' }
    });

    const evenRowStyle = Style({
        backgroundColor: '#f3f3f3'
    });

    // 3. Define the entire sheet layout using components
    const myToDoList = VStack({
        // Base style for the entire table
        style: Style({ font: { family: 'Arial', size: 10 } }),
        children: [
            // Header Row
            HStack({
                style: headerStyle,
                children: [
                    Cell({ type: Text('TASK ID'), colSpan: 2 }),
                    Cell({ type: Text('DESCRIPTION'), colSpan: 3 }),
                    Cell({ type: Text('STATUS') }),
                    Cell({ type: Text('DUE DATE') }),
                    Cell({ type: Text('DONE') })
                ]
            }),
            // Data Row 1
            HStack({
                children: [
                    Cell({ type: Text('PROJ-A'), style: Style({ font: { bold: true } }) }),
                    Cell({ type: Text('101') }),
                    Cell({ type: Text('Finalize Q3 report and submit for review.'), colSpan: 3 }),
                    Cell({
                        type: Dropdown({
                            values: ['Pending', 'In Progress', 'Complete'],
                            selected: 'In Progress'
                        })
                    }),
                    Cell({ type: DatePicker({ format: 'yyyy-mm-dd' }) }),
                    Cell({ type: Checkbox(false) })
                ]
            }),
            // Data Row 2
            HStack({
                style: evenRowStyle, // Apply a style to the whole row
                children: [
                    Cell({ type: Text('PROJ-B'), style: Style({ font: { bold: true } }) }),
                    Cell({ type: Text('205') }),
                    Cell({ type: Text('Onboard new team members.'), colSpan: 3 }),
                    Cell({
                        type: Dropdown({
                            values: ['Pending', 'In Progress', 'Complete'],
                            selected: 'Pending'
                        })
                    }),
                    Cell({ type: DatePicker({ format: 'yyyy-mm-dd' }) }),
                    Cell({ type: Checkbox(false) })
                ]
            }),
             // Data Row 3
            HStack({
                children: [
                    Cell({ type: Text('PROJ-A'), style: Style({ font: { bold: true } }) }),
                    Cell({ type: Text('102') }),
                    Cell({ type: Text('Prepare slides for stakeholder meeting.'), colSpan: 3 }),
                    Cell({
                        type: Dropdown({
                            values: ['Pending', 'In Progress', 'Complete'],
                            selected: 'Complete'
                        })
                    }),
                    Cell({ type: DatePicker({ format: 'yyyy-mm-dd' }) }),
                    Cell({ type: Checkbox(true) })
                ]
            }),
        ]
    });

    // 4. Render the final layout to the sheet
    render(myToDoList, sheet);
    
    // Adjust column widths for better readability
    sheet.autoResizeColumns(1, sheet.getLastColumn());
}
