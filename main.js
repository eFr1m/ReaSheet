/**
 * =======================================================================
 * ReaSheets Examples
 *
 * This file contains practical examples demonstrating various features of
 * the ReaSheets framework, including cell spanning, collision handling,
 * complex layouts, forms, and conditional formatting.
 *
 * To run these examples in Google Apps Script:
 * 1. Select one of the example functions (e.g., cellSpanningExample)
 * 2. Click the "Run" button in the Apps Script editor
 * =======================================================================
 */

/**
 * Example 1: Demonstrates cell spanning and how the occupancyMap prevents collisions.
 *
 * When cells use colSpan or rowSpan, the layout engine tracks occupied positions
 * and automatically advances the cursor to avoid overlaps.
 */
function cellSpanningExample() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName("CellSpanning") || ss.insertSheet("CellSpanning");
  sheet.clear();

  const headerStyle = new Style({
    backgroundColor: "#1f4e78",
    font: { color: "white", bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    height: 35,
  });

  const layout = new VStack({
    children: [
      new HStack({
        style: headerStyle,
        children: [
          new Cell({
            type: new Text("Cell Spanning & Collision Handling Demo"),
            colSpan: 4,
          }),
        ],
      }),

      new Cell({ type: new Text("") }), // Spacer

      // Example 1: Simple column spans
      new HStack({
        children: [
          new Cell({
            type: new Text("1 col"),
            style: new Style({
              backgroundColor: "#cfe2f3",
              alignment: { horizontal: "center" },
            }),
          }),
          new Cell({
            type: new Text("2 cols"),
            colSpan: 2,
            style: new Style({
              backgroundColor: "#fff2cc",
              alignment: { horizontal: "center" },
            }),
          }),
          new Cell({
            type: new Text("1 col"),
            style: new Style({
              backgroundColor: "#d9ead3",
              alignment: { horizontal: "center" },
            }),
          }),
        ],
      }),

      // Example 2: Row and column spans combined
      new HStack({
        children: [
          new Cell({
            type: new Text("A"),
            rowSpan: 2,
            style: new Style({
              backgroundColor: "#f8cbad",
              alignment: { horizontal: "center", vertical: "middle" },
              height: 50,
            }),
          }),
          new Cell({
            type: new Text("B"),
            style: new Style({
              backgroundColor: "#ea9999",
              alignment: { horizontal: "center" },
            }),
          }),
          new Cell({
            type: new Text("C - Spans 2 cols"),
            colSpan: 2,
            style: new Style({
              backgroundColor: "#e2efda",
              alignment: { horizontal: "center" },
            }),
          }),
        ],
      }),

      // Next row - the occupancyMap prevents collision with "A"
      new HStack({
        children: [
          new Cell({
            type: new Text("D"),
            style: new Style({
              backgroundColor: "#c9daf8",
              alignment: { horizontal: "center" },
            }),
          }),
          new Cell({
            type: new Text("E"),
            style: new Style({
              backgroundColor: "#f4cccc",
              alignment: { horizontal: "center" },
            }),
          }),
          new Cell({
            type: new Text("F"),
            style: new Style({
              backgroundColor: "#d5a6bd",
              alignment: { horizontal: "center" },
            }),
          }),
        ],
      }),

      new Cell({ type: new Text("") }), // Spacer

      new Cell({
        type: new Text(
          "How it works: The occupancyMap tracks cells A1, A2 (from 'A' with rowSpan: 2). When the second row is laid out, it automatically skips to column 2 to avoid the collision."
        ),
        style: new Style({
          font: { italic: true, size: 9 },
          wrap: { strategy: WrapStrategy.WRAP },
        }),
      }),
    ],
  });

  render(layout, sheet);
}

/**
 * Example 2: A complex dashboard with multiple sections and mixed layouts.
 *
 * Demonstrates nested VStack and HStack components with varied column widths.
 */
function dashboardExample() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  sheet.clear();
  sheet.setFrozenRows(1);

  const headerStyle = new Style({
    backgroundColor: "#1f4e78",
    font: { color: "white", bold: true, size: 14 },
    alignment: { horizontal: "center", vertical: "middle" },
    height: 40,
  });

  const sectionHeaderStyle = new Style({
    backgroundColor: "#d9e1f2",
    font: { bold: true },
    border: new Border({
      bottom: { color: "#1f4e78", thickness: BorderThickness.SOLID },
    }),
  });

  const metricStyle = new Style({
    backgroundColor: "#f2f2f2",
    alignment: { horizontal: "right" },
  });

  const dashboard = new VStack({
    children: [
      // Main header
      new HStack({
        style: headerStyle,
        children: [
          new Cell({ type: new Text("Q4 2024 Sales Dashboard"), colSpan: 6 }),
        ],
      }),

      // Metrics section
      new Cell({
        type: new Text("KEY METRICS"),
        style: sectionHeaderStyle,
      }),

      new HStack({
        children: [
          new Cell({
            type: new Text("Total Revenue:"),
            style: new Style({ font: { bold: true } }),
          }),
          new Cell({
            type: new Number(250000, NumberFormats.CURRENCY),
            style: metricStyle,
          }),
          new Cell({ type: new Text("") }),
          new Cell({
            type: new Text("Growth Rate:"),
            style: new Style({ font: { bold: true } }),
          }),
          new Cell({
            type: new Number(0.15, NumberFormats.PERCENTAGE),
            style: metricStyle,
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({
            type: new Text("Active Customers:"),
            style: new Style({ font: { bold: true } }),
          }),
          new Cell({
            type: new Number(1250, "0"),
            style: metricStyle,
          }),
          new Cell({ type: new Text("") }),
          new Cell({
            type: new Text("Conversion:"),
            style: new Style({ font: { bold: true } }),
          }),
          new Cell({
            type: new Number(0.32, NumberFormats.PERCENTAGE),
            style: metricStyle,
          }),
        ],
      }),

      new Cell({ type: new Text("") }), // Spacer

      // Regional breakdown
      new Cell({
        type: new Text("REGIONAL PERFORMANCE"),
        style: sectionHeaderStyle,
      }),

      new HStack({
        style: new Style({ backgroundColor: "#ecf0f1" }),
        children: [
          new Cell({
            type: new Text("Region"),
            style: new Style({ font: { bold: true } }),
          }),
          new Cell({
            type: new Text("Sales"),
            style: new Style({ font: { bold: true } }),
          }),
          new Cell({
            type: new Text("Units"),
            style: new Style({ font: { bold: true } }),
          }),
          new Cell({
            type: new Text("Avg Price"),
            style: new Style({ font: { bold: true } }),
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("North America") }),
          new Cell({ type: new Number(120000, NumberFormats.CURRENCY) }),
          new Cell({ type: new Number(450, "0") }),
          new Cell({ type: new Number(266.67, NumberFormats.CURRENCY) }),
        ],
      }),

      new HStack({
        style: new Style({ backgroundColor: "#f2f2f2" }),
        children: [
          new Cell({ type: new Text("Europe") }),
          new Cell({ type: new Number(85000, NumberFormats.CURRENCY) }),
          new Cell({ type: new Number(320, "0") }),
          new Cell({ type: new Number(265.63, NumberFormats.CURRENCY) }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Asia Pacific") }),
          new Cell({ type: new Number(45000, NumberFormats.CURRENCY) }),
          new Cell({ type: new Number(180, "0") }),
          new Cell({ type: new Number(250, NumberFormats.CURRENCY) }),
        ],
      }),
    ],
  });

  render(dashboard, sheet);
}

/**
 * Example 3: A data entry form with various input types and validation.
 *
 * Demonstrates dropdowns, date pickers, checkboxes, and form layout patterns.
 */
function dataEntryFormExample() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DataForm") || ss.insertSheet("DataForm");
  sheet.clear();

  const labelStyle = new Style({
    font: { bold: true },
    alignment: { horizontal: "right", vertical: "middle" },
    width: 150,
  });

  const inputCellStyle = new Style({
    backgroundColor: "#ffffff",
    border: new Border({
      bottom: { color: "#cccccc", thickness: BorderThickness.SOLID },
    }),
    width: 300,
  });

  const sectionHeaderStyle = new Style({
    font: { bold: true, size: 12 },
    backgroundColor: "#d9e1f2",
    border: new Border({
      bottom: { color: "black", thickness: BorderThickness.SOLID },
    }),
  });

  const form = new VStack({
    style: new Style({ font: { family: "Arial", size: 11 } }),
    children: [
      new Cell({
        type: new Text("Customer Information Form"),
        style: new Style({
          font: { bold: true, size: 16 },
          height: 40,
          alignment: { horizontal: "center" },
        }),
      }),

      new Cell({ type: new Text("") }), // Spacer

      new Cell({
        type: new Text("PERSONAL INFORMATION"),
        style: sectionHeaderStyle,
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Full Name:"), style: labelStyle }),
          new Cell({ type: new Text(""), style: inputCellStyle, colSpan: 2 }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Email:"), style: labelStyle }),
          new Cell({ type: new Text(""), style: inputCellStyle, colSpan: 2 }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Phone:"), style: labelStyle }),
          new Cell({ type: new Text(""), style: inputCellStyle, colSpan: 2 }),
        ],
      }),

      new Cell({ type: new Text("") }), // Spacer

      new Cell({
        type: new Text("ACCOUNT DETAILS"),
        style: sectionHeaderStyle,
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Customer Type:"), style: labelStyle }),
          new Cell({
            type: new Dropdown({
              values: [
                {
                  value: "Individual",
                  style: new Style({ backgroundColor: "#fff2cc" }),
                },
                {
                  value: "Business",
                  style: new Style({ backgroundColor: "#cfe2f3" }),
                },
                {
                  value: "Non-Profit",
                  style: new Style({ backgroundColor: "#d9ead3" }),
                },
              ],
              selected: "Individual",
            }),
            style: inputCellStyle,
            colSpan: 2,
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Registration Date:"), style: labelStyle }),
          new Cell({
            type: new DatePicker({ format: "yyyy-mm-dd" }),
            style: inputCellStyle,
            colSpan: 2,
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Newsletter:"), style: labelStyle }),
          new Cell({
            type: new Checkbox(false),
            style: new Style({ alignment: { horizontal: "center" } }),
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Premium Member:"), style: labelStyle }),
          new Cell({
            type: new Checkbox(false),
            style: new Style({ alignment: { horizontal: "center" } }),
          }),
        ],
      }),
    ],
  });

  render(form, sheet);
}

/**
 * Example 4: A sales table with conditional formatting based on status.
 *
 * Demonstrates how the Dropdown component can have conditional formatting rules
 * that apply different styles based on selected values.
 */
function salesTableExample() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SalesTable") || ss.insertSheet("SalesTable");
  sheet.clear();
  sheet.setFrozenRows(1);

  const headerStyle = new Style({
    backgroundColor: "#34495e",
    font: { color: "white", bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    height: 35,
  });

  const data = [
    { product: "Laptop", q1: 15000, q2: 18000, q3: 21000, status: "Complete" },
    { product: "Tablet", q1: 8000, q2: 9500, q3: 11200, status: "In Progress" },
    { product: "Phone", q1: 25000, q2: 28000, q3: 31000, status: "Complete" },
    { product: "Monitor", q1: 3500, q2: 4200, q3: 5100, status: "Pending" },
    {
      product: "Keyboard",
      q1: 2200,
      q2: 2800,
      q3: 3500,
      status: "In Progress",
    },
  ];

  const statusOptions = [
    {
      value: "Pending",
      style: new Style({
        backgroundColor: "#fff2cc",
        font: { color: "#8d6e63" },
      }),
    },
    {
      value: "In Progress",
      style: new Style({
        backgroundColor: "#cfe2f3",
        font: { color: "#1a237e" },
      }),
    },
    {
      value: "Complete",
      style: new Style({
        backgroundColor: "#d9ead3",
        font: { color: "#1b5e20" },
      }),
    },
  ];

  const tableLayout = new VStack({
    style: new Style({ font: { family: "Arial", size: 10 } }),
    children: [
      // Header
      new HStack({
        style: headerStyle,
        children: [
          new Cell({ type: new Text("Product") }),
          new Cell({ type: new Text("Q1 Sales") }),
          new Cell({ type: new Text("Q2 Sales") }),
          new Cell({ type: new Text("Q3 Sales") }),
          new Cell({ type: new Text("Status") }),
        ],
      }),

      // Data rows with alternating backgrounds
      ...data.map(
        (row, idx) =>
          new HStack({
            style:
              idx % 2 === 0
                ? new Style({ backgroundColor: "#f8f9fa" })
                : new Style(),
            children: [
              new Cell({ type: new Text(row.product) }),
              new Cell({ type: new Number(row.q1, NumberFormats.CURRENCY) }),
              new Cell({ type: new Number(row.q2, NumberFormats.CURRENCY) }),
              new Cell({ type: new Number(row.q3, NumberFormats.CURRENCY) }),
              new Cell({
                type: new Dropdown({
                  values: statusOptions,
                  selected: row.status,
                }),
              }),
            ],
          })
      ),

      new Cell({ type: new Text("") }), // Spacer

      new Cell({
        type: new Text(
          "Tip: The status column uses conditional formatting. Try changing a status to see the color update."
        ),
        style: new Style({
          font: { italic: true, size: 9 },
          wrap: { strategy: WrapStrategy.WRAP },
        }),
      }),
    ],
  });

  render(tableLayout, sheet);
}

/**
 * Example 5: A feature showcase demonstrating all styling options.
 *
 * This example shows off borders, rotations, alignment, wrapping, and more.
 */
function stylingSampleExample() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName("StylingDemo") || ss.insertSheet("StylingDemo");
  sheet.clear();

  const TestSection = ({ title, children }) =>
    new VStack({
      children: [
        new Cell({
          type: new Text(title),
          style: new Style({
            font: { bold: true, size: 12 },
            backgroundColor: "#e8eaf6",
          }),
        }),
        ...children,
      ],
    });

  const layout = new VStack({
    children: [
      new Cell({
        type: new Text("ReaSheets Styling Demo"),
        style: new Style({
          font: { bold: true, size: 18 },
          height: 40,
          alignment: { horizontal: "center" },
        }),
      }),

      TestSection({
        title: "Font Styles",
        children: [
          new HStack({
            children: [
              new Cell({
                type: new Text("Bold"),
                style: new Style({ font: { bold: true } }),
              }),
              new Cell({
                type: new Text("Italic"),
                style: new Style({ font: { italic: true } }),
              }),
              new Cell({
                type: new Text("Underline"),
                style: new Style({ font: { underline: true } }),
              }),
              new Cell({
                type: new Text("Strikethrough"),
                style: new Style({ font: { strikethrough: true } }),
              }),
            ],
          }),
        ],
      }),

      TestSection({
        title: "Colors & Backgrounds",
        children: [
          new HStack({
            children: [
              new Cell({
                type: new Text("Red Text"),
                style: new Style({ font: { color: "#ff0000" } }),
              }),
              new Cell({
                type: new Text("Blue BG"),
                style: new Style({ backgroundColor: "#cfe2f3" }),
              }),
              new Cell({
                type: new Text("Green Text + Yellow BG"),
                style: new Style({
                  backgroundColor: "#fff2cc",
                  font: { color: "#38761d" },
                }),
              }),
            ],
          }),
        ],
      }),

      TestSection({
        title: "Alignment & Wrapping",
        children: [
          new HStack({
            style: new Style({ height: 60 }),
            children: [
              new Cell({
                type: new Text("Left Top"),
                style: new Style({
                  alignment: { horizontal: "left", vertical: "top" },
                  width: 120,
                }),
              }),
              new Cell({
                type: new Text("Center Middle"),
                style: new Style({
                  alignment: { horizontal: "center", vertical: "middle" },
                  width: 120,
                }),
              }),
              new Cell({
                type: new Text("Right Bottom"),
                style: new Style({
                  alignment: { horizontal: "right", vertical: "bottom" },
                  width: 120,
                }),
              }),
            ],
          }),

          new HStack({
            children: [
              new Cell({
                type: new Text(
                  "This is a long text that demonstrates the WRAP strategy. It will wrap within the cell bounds."
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

      TestSection({
        title: "Borders",
        children: [
          new HStack({
            children: [
              new Cell({
                type: new Text("Solid Border"),
                style: new Style({
                  border: new Border({
                    top: { color: "black", thickness: BorderThickness.SOLID },
                    bottom: {
                      color: "black",
                      thickness: BorderThickness.SOLID,
                    },
                    left: { color: "black", thickness: BorderThickness.SOLID },
                    right: { color: "black", thickness: BorderThickness.SOLID },
                  }),
                }),
              }),
              new Cell({
                type: new Text("Dashed Border"),
                style: new Style({
                  border: new Border({
                    top: {
                      color: "#2196f3",
                      thickness: BorderThickness.DASHED,
                    },
                    bottom: {
                      color: "#2196f3",
                      thickness: BorderThickness.DASHED,
                    },
                    left: {
                      color: "#2196f3",
                      thickness: BorderThickness.DASHED,
                    },
                    right: {
                      color: "#2196f3",
                      thickness: BorderThickness.DASHED,
                    },
                  }),
                }),
              }),
              new Cell({
                type: new Text("Thick Border"),
                style: new Style({
                  border: new Border({
                    top: {
                      color: "#ff0000",
                      thickness: BorderThickness.SOLID_THICK,
                    },
                    bottom: {
                      color: "#ff0000",
                      thickness: BorderThickness.SOLID_THICK,
                    },
                    left: {
                      color: "#ff0000",
                      thickness: BorderThickness.SOLID_THICK,
                    },
                    right: {
                      color: "#ff0000",
                      thickness: BorderThickness.SOLID_THICK,
                    },
                  }),
                }),
              }),
            ],
          }),
        ],
      }),

      TestSection({
        title: "Numbers & Validation",
        children: [
          new HStack({
            children: [
              new Cell({
                type: new Number(1234.56, NumberFormats.CURRENCY),
                style: new Style({ alignment: { horizontal: "right" } }),
              }),
              new Cell({
                type: new Number(0.75, NumberFormats.PERCENTAGE),
                style: new Style({ alignment: { horizontal: "right" } }),
              }),
              new Cell({
                type: new Number(123.456, "0.00"),
                style: new Style({ alignment: { horizontal: "right" } }),
              }),
            ],
          }),

          new HStack({
            children: [
              new Cell({
                type: new Checkbox(true),
                style: new Style({ alignment: { horizontal: "center" } }),
              }),
              new Cell({
                type: new Checkbox(false),
                style: new Style({ alignment: { horizontal: "center" } }),
              }),
              new Cell({
                type: new DatePicker({ format: "yyyy-mm-dd" }),
                style: new Style({ alignment: { horizontal: "center" } }),
              }),
            ],
          }),
        ],
      }),
    ],
  });

  render(layout, sheet);
}

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
