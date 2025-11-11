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

function ToDoItem({ id, description, status, isDone, isEven = false }) {
  const rowStyle = isEven
    ? new Style({ backgroundColor: "#f3f3f3" })
    : new Style();

  // Dropdown options with conditional formatting
  const statusOptions = [
    { value: "Pending", style: new Style({ backgroundColor: "#fff2cc" }) },
    { value: "In Progress", style: new Style({ backgroundColor: "#cfe2f3" }) },
    { value: "Complete", style: new Style({ backgroundColor: "#d9ead3" }) },
  ];

  return new HStack({
    style: rowStyle,
    children: [
      new Cell({
        type: new Text(id.split("-")[0]),
        style: new Style({ font: { bold: true } }),
      }),
      new Cell({ type: new Text(id.split("-")[1]) }),
      new Cell({ type: new Text(description), colSpan: 3 }),
      new Cell({
        type: new Dropdown({
          values: statusOptions,
          selected: status,
        }),
      }),
      new Cell({
        type: new DatePicker({ format: "yyyy-mm-dd", value: new Date() }),
      }),
      new Cell({ type: new Checkbox(isDone) }),
    ],
  });
}

/**
 * The main function to generate the sheet.
 * This function gets the target sheet, defines the layout using components,
 * and then calls the render function to build the sheet.
 */
function createToDoListSheet() {
  // 1. Get the target sheet object from Google Apps Script
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ToDoList") || ss.insertSheet("ToDoList");

  // It's good practice to clear the sheet for a fresh render
  sheet.clear();
  sheet.setFrozenRows(1);

  // 2. Define some reusable styles for clarity
  const headerStyle = new Style({
    backgroundColor: "#4a86e8",
    font: { color: "white", bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    border: new Border({
      bottom: { color: "black", thickness: BorderThickness.SOLID_THICK },
    }),
    height: 40,
  });

  const header = new HStack({
    style: headerStyle,
    children: [
      new Cell({ type: new Text("TASK ID"), colSpan: 2 }),
      new Cell({ type: new Text("DESCRIPTION"), colSpan: 3 }),
      new Cell({ type: new Text("STATUS") }),
      new Cell({ type: new Text("DUE DATE") }),
      new Cell({ type: new Text("DONE") }),
    ],
  });

  // 3. Define the entire sheet layout using components
  const myToDoList = new VStack({
    // Base style for the entire table
    style: new Style({ font: { family: "Arial", size: 10 } }),
    children: [
      header,
      ToDoItem({
        id: "PROJ-A-101",
        description: "Finalize Q3 report.",
        status: "In Progress",
        isDone: false,
      }),
      ToDoItem({
        id: "PROJ-B-205",
        description: "Onboard new team members.",
        status: "Pending",
        isDone: false,
        isEven: true,
      }),
      ToDoItem({
        id: "PROJ-A-102",
        description: "Prepare slides for stakeholder meeting.",
        status: "Complete",
        isDone: true,
      }),
    ],
  });

  // 4. Render the final layout to the sheet
  render(myToDoList, sheet);
}

function createFeatureTestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName("FeatureTest") || ss.insertSheet("FeatureTest");
  sheet.clear();

  const TestSection = ({ title, children }) =>
    new VStack({
      children: [
        new Cell({
          type: new Text(title),
          style: new Style({ font: { bold: true, size: 14 } }),
        }),
        ...children,
      ],
    });

  const testLayout = new VStack({
    children: [
      new Cell({
        type: new Text("Feature Test Sheet"),
        style: new Style({ font: { bold: true, size: 18 } }),
      }),

      TestSection({
        title: "--- Text Tests ---",
        children: [
          new HStack({
            children: [
              new Cell({ type: new Text("Regular Text") }),
              new Cell({ type: new Text("12345") }),
              new Cell({ type: new Text("123.45") }),
              new Cell({ type: new Text("-50") }),
            ],
          }),
        ],
      }),

      TestSection({
        title: "--- Number Tests ---",
        children: [
          new HStack({
            children: [
              new Cell({ type: new Number(123) }),
              new Cell({ type: new Number(0.98, NumberFormats.PERCENTAGE) }),
              new Cell({ type: new Number(1234.56, NumberFormats.CURRENCY) }),
              new Cell({ type: new Number(123.456, "0.00") }),
            ],
          }),
        ],
      }),

      TestSection({
        title: "--- Style Tests ---",
        children: [
          new HStack({
            style: new Style({ height: 50 }),
            children: [
              new Cell({
                type: new Text("BG Color"),
                style: new Style({ backgroundColor: "#cfe2f3" }),
              }),
              new Cell({
                type: new Text("Font Style"),
                style: new Style({
                  font: { color: "red", italic: true, strikethrough: true },
                }),
              }),
              new Cell({
                type: new Text("Alignment"),
                style: new Style({
                  alignment: { horizontal: "center", vertical: "middle" },
                }),
              }),
              new Cell({
                type: new Text("Rotation"),
                style: new Style({ rotation: { angle: 45 } }),
              }),
            ],
          }),
        ],
      }),

      TestSection({
        title: "--- Width/Height & Border Tests ---",
        children: [
          new HStack({
            style: new Style({ height: 60 }),
            children: [
              new Cell({
                type: new Text("Width 200"),
                style: new Style({
                  width: 200,
                  border: new Border({
                    right: {
                      color: "red",
                      thickness: BorderThickness.SOLID_THICK,
                    },
                  }),
                }),
              }),
              new Cell({
                type: new Text(
                  "This cell should inherit the height of 60 from the parent HStack."
                ),
                style: new Style({
                  border: new Border({
                    left: { color: "blue", thickness: BorderThickness.DASHED },
                  }),
                }),
              }),
            ],
          }),
        ],
      }),

      TestSection({
        title: "--- Component Tests ---",
        children: [
          new HStack({
            children: [
              new Cell({ type: new Checkbox(true) }),
              new Cell({ type: new Checkbox(false) }),
              new Cell({
                type: new Dropdown({
                  values: ["A", "B", "C"],
                  selected: "B",
                }),
              }),
              new Cell({
                type: new DatePicker({
                  value: new Date(),
                  format: "yyyy-mm-dd",
                }),
              }),
            ],
          }),
        ],
      }),

      TestSection({
        title: "--- Wrap Strategy Test ---",
        children: [
          new HStack({
            children: [
              new Cell({
                type: new Text(
                  "This is a long text that should wrap inside the cell."
                ),
                style: new Style({
                  width: 150,
                  wrap: { strategy: WrapStrategy.WRAP },
                }),
              }),
            ],
          }),
        ],
      }),
    ],
  });

  render(testLayout, sheet);
}
