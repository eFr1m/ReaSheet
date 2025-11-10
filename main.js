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
    const headerStyle = new Style({
        backgroundColor: '#4a86e8',
        font: { color: 'white', bold: true },
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: new Border({
            bottom: { color: 'black', thickness: BorderThickness.SOLID_THICK }
        })
    });

    const evenRowStyle = new Style({
        backgroundColor: '#f3f3f3'
    });

    // Define styles for dropdown options
    const pendingStyle = new Style({ backgroundColor: '#fff2cc' }); // Light orange
    const inProgressStyle = new Style({ backgroundColor: '#cfe2f3' }); // Light blue
    const completeStyle = new Style({ backgroundColor: '#d9ead3' }); // Light green

    // 3. Define the entire sheet layout using components
    const myToDoList = new VStack({
        // Base style for the entire table
        style: new Style({ font: { family: 'Arial', size: 10 } }),
        children: [
            // Header Row
            new HStack({
                style: headerStyle,
                children: [
                    new Cell({ type: new Text('TASK ID'), colSpan: 2 }),
                    new Cell({ type: new Text('DESCRIPTION'), colSpan: 3 }),
                    new Cell({ type: new Text('STATUS') }),
                    new Cell({ type: new Text('DUE DATE') }),
                    new Cell({ type: new Text('DONE') })
                ]
            }),
            // Data Row 1
            new HStack({
                children: [
                    new Cell({ type: new Text('PROJ-A'), style: new Style({ font: { bold: true } }) }),
                    new Cell({ type: new Text('101') }),
                    new Cell({ type: new Text('Finalize Q3 report and submit for review.'), colSpan: 3 }),
                    new Cell({
                        type: new Dropdown({
                            values: [
                                { value: 'Pending', style: pendingStyle },
                                { value: 'In Progress', style: inProgressStyle },
                                { value: 'Complete', style: completeStyle }
                            ],
                            selected: 'In Progress'
                        })
                    }),
                    new Cell({ type: new DatePicker({ format: 'yyyy-mm-dd' }) }),
                    new Cell({ type: new Checkbox(false) })
                ]
            }),
            // Data Row 2
            new HStack({
                style: evenRowStyle, // Apply a style to the whole row
                children: [
                    new Cell({ type: new Text('PROJ-B'), style: new Style({ font: { bold: true } }) }),
                    new Cell({ type: new Text('205') }),
                    new Cell({ type: new Text('Onboard new team members.'), colSpan: 3 }),
                    new Cell({
                        type: new Dropdown({
                            values: [
                                { value: 'Pending', style: pendingStyle },
                                { value: 'In Progress', style: inProgressStyle },
                                { value: 'Complete', style: completeStyle }
                            ],
                            selected: 'Pending'
                        })
                    }),
                    new Cell({ type: new DatePicker({ format: 'yyyy-mm-dd' }) }),
                    new Cell({ type: new Checkbox(false) })
                ]
            }),
             // Data Row 3
            new HStack({
                children: [
                    new Cell({ type: new Text('PROJ-A'), style: new Style({ font: { bold: true } }) }),
                    new Cell({ type: new Text('102') }),
                    new Cell({ type: new Text('Prepare slides for stakeholder meeting.'), colSpan: 3 }),
                    new Cell({
                        type: new Dropdown({
                            values: [
                                { value: 'Pending', style: pendingStyle },
                                { value: 'In Progress', style: inProgressStyle },
                                { value: 'Complete', style: completeStyle }
                            ],
                            selected: 'Complete'
                        })
                    }),
                    new Cell({ type: new DatePicker({ format: 'yyyy-mm-dd' }) }),
                    new Cell({ type: new Checkbox(true) })
                ]
            }),
        ]
    });

    // 4. Render the final layout to the sheet
    render(myToDoList, sheet);
    
    // Adjust column widths for better readability
    sheet.autoResizeColumns(1, sheet.getLastColumn());
}
