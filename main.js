/**
 * =======================================================================
 * ReaSheets Examples
 *
 * 1) `basicProductCard`      - simple product display with styling
 * 2) `taskTrackerDemo`       - interactive task manager with conditional formatting
 * 3) `dashboardExample`      - complex dashboard with KPIs and data tables
 * =======================================================================
 */

function basicProductCard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName("ProductCard") || ss.insertSheet("ProductCard");
  sheet.clear();

  const layout = new VStack({
    children: [
      new Cell({
        type: new Text("Product Showcase"),
        style: new Style({
          font: { bold: true, size: 18, color: "white" },
          backgroundColor: "#4285f4",
          alignment: { horizontal: "center", vertical: "middle" },
          height: 50,
        }),
        colSpan: 2,
      }),

      new HStack({
        children: [
          new Cell({
            type: new Text("Product:"),
            style: new Style({
              font: { bold: true },
              alignment: { horizontal: "right" },
              width: 120,
            }),
          }),
          new Cell({
            type: new Text("Wireless Headphones"),
            style: new Style({ width: 200 }),
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({
            type: new Text("Price:"),
            style: new Style({
              font: { bold: true },
              alignment: { horizontal: "right" },
            }),
          }),
          new Cell({
            type: new NumberCell(79.99, NumberFormat.CURRENCY),
            style: new Style({
              font: { size: 12, color: "#0d652d" },
              backgroundColor: "#d9ead3",
            }),
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({
            type: new Text("In Stock:"),
            style: new Style({
              font: { bold: true },
              alignment: { horizontal: "right" },
            }),
          }),
          new Cell({
            type: new Checkbox(true),
            style: new Style({ alignment: { horizontal: "center" } }),
          }),
        ],
      }),
    ],
  });

  render(sheet, layout);
}

function taskTrackerDemo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName("TaskTracker") || ss.insertSheet("TaskTracker");
  sheet.clear();
  sheet.setFrozenRows(1);

  const headerStyle = new Style({
    backgroundColor: "#34a853",
    font: { color: "white", bold: true, size: 12 },
    alignment: { horizontal: "center", vertical: "middle" },
    height: 40,
  });

  const statusOptions = [
    { value: "To Do", style: new Style({ backgroundColor: "#fce5cd" }) },
    { value: "In Progress", style: new Style({ backgroundColor: "#cfe2f3" }) },
    { value: "Done", style: new Style({ backgroundColor: "#d9ead3" }) },
    { value: "Blocked", style: new Style({ backgroundColor: "#f4cccc" }) },
  ];

  const priorityOptions = [
    { value: "Low", style: new Style({ backgroundColor: "#d9ead3" }) },
    { value: "Medium", style: new Style({ backgroundColor: "#fff2cc" }) },
    { value: "High", style: new Style({ backgroundColor: "#f4cccc" }) },
  ];

  const layout = new VStack({
    children: [
      new HStack({
        style: headerStyle,
        children: [
          new Cell({ type: new Text("Done"), colSpan: 1 }),
          new Cell({ type: new Text("Task"), colSpan: 1 }),
          new Cell({ type: new Text("Status"), colSpan: 1 }),
          new Cell({ type: new Text("Priority"), colSpan: 1 }),
          new Cell({ type: new Text("Due Date"), colSpan: 1 }),
        ],
      }),

      new HStack({
        children: [
          new Cell({
            type: new Checkbox(false),
            style: new Style({
              alignment: { horizontal: "center" },
              width: 60,
            }),
          }),
          new Cell({
            type: new Text("Implement user authentication"),
            style: new Style({ width: 250 }),
          }),
          new Cell({
            type: new Dropdown({
              values: statusOptions,
              selected: "In Progress",
            }),
            style: new Style({ width: 120 }),
          }),
          new Cell({
            type: new Dropdown({ values: priorityOptions, selected: "High" }),
            style: new Style({ width: 100 }),
          }),
          new Cell({
            type: new DatePicker({ format: "yyyy-mm-dd" }),
            style: new Style({ width: 120 }),
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
            type: new Text("Setup database"),
            style: new Style({
              font: { strikethrough: true, color: "#999999" },
            }),
          }),
          new Cell({
            type: new Dropdown({ values: statusOptions, selected: "Done" }),
          }),
          new Cell({
            type: new Dropdown({ values: priorityOptions, selected: "Medium" }),
          }),
          new Cell({ type: new DatePicker(new Date("2025-12-20")) }),
        ],
      }),

      new HStack({
        children: [
          new Cell({
            type: new Checkbox(false),
            style: new Style({ alignment: { horizontal: "center" } }),
          }),
          new Cell({ type: new Text("Write API documentation") }),
          new Cell({
            type: new Dropdown({ values: statusOptions, selected: "To Do" }),
          }),
          new Cell({
            type: new Dropdown({ values: priorityOptions, selected: "Low" }),
          }),
          new Cell({ type: new DatePicker({ format: "yyyy-mm-dd" }) }),
        ],
      }),
    ],
  });

  render(sheet, layout);
}

function dashboardExample() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  sheet.clear();
  sheet.setFrozenRows(1);

  const titleStyle = new Style({
    backgroundColor: "#1a73e8",
    font: { color: "white", bold: true, size: 16 },
    alignment: { horizontal: "center", vertical: "middle" },
    height: 50,
  });

  const kpiLabelStyle = new Style({
    font: { bold: true, size: 11 },
    alignment: { horizontal: "center", vertical: "bottom" },
    height: 30,
  });

  const kpiValueStyle = new Style({
    font: { bold: true, size: 18 },
    alignment: { horizontal: "center", vertical: "middle" },
    height: 40,
  });

  const tableHeaderStyle = new Style({
    backgroundColor: "#666666",
    font: { color: "white", bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    height: 30,
  });

  const statusOptions = [
    { value: "Active", style: new Style({ backgroundColor: "#d9ead3" }) },
    { value: "Pending", style: new Style({ backgroundColor: "#fff2cc" }) },
    { value: "Inactive", style: new Style({ backgroundColor: "#f4cccc" }) },
  ];

  const dashboard = new VStack({
    children: [
      new HStack({
        style: titleStyle,
        children: [
          new Cell({
            type: new Text("📊 Sales Dashboard - Q4 2025"),
            colSpan: 6,
          }),
        ],
      }),

      new HStack({
        children: [
          new VStack({
            children: [
              new Cell({
                type: new Text("Total Revenue"),
                style: kpiLabelStyle,
              }),
              new Cell({
                type: new NumberCell(458920, NumberFormat.CURRENCY),
                style: {
                  ...kpiValueStyle,
                  backgroundColor: "#d9ead3",
                  font: { ...kpiValueStyle.font, color: "#0d652d" },
                },
              }),
            ],
          }),
          new VStack({
            children: [
              new Cell({ type: new Text("Orders"), style: kpiLabelStyle }),
              new Cell({
                type: new NumberCell(1247, NumberFormat.INTEGER),
                style: {
                  ...kpiValueStyle,
                  backgroundColor: "#cfe2f3",
                  font: { ...kpiValueStyle.font, color: "#1c4587" },
                },
              }),
            ],
          }),
          new VStack({
            children: [
              new Cell({
                type: new Text("Avg Order Value"),
                style: kpiLabelStyle,
              }),
              new Cell({
                type: new NumberCell(368, NumberFormat.CURRENCY),
                style: {
                  ...kpiValueStyle,
                  backgroundColor: "#fff2cc",
                  font: { ...kpiValueStyle.font, color: "#7f6000" },
                },
              }),
            ],
          }),
        ],
      }),

      new Cell({
        type: new Text(""),
        style: new Style({ height: 20 }),
        colSpan: 6,
      }),

      new HStack({
        style: tableHeaderStyle,
        children: [
          new Cell({ type: new Text("Product"), colSpan: 1 }),
          new Cell({ type: new Text("Status"), colSpan: 1 }),
          new Cell({ type: new Text("Units Sold"), colSpan: 1 }),
          new Cell({ type: new Text("Revenue"), colSpan: 1 }),
          new Cell({ type: new Text("Growth %"), colSpan: 1 }),
          new Cell({ type: new Text("Target Met"), colSpan: 1 }),
        ],
      }),

      new HStack({
        children: [
          new Cell({
            type: new Text("Laptop Pro"),
            style: new Style({ width: 140 }),
          }),
          new Cell({
            type: new Dropdown({ values: statusOptions, selected: "Active" }),
            style: new Style({ width: 100 }),
          }),
          new Cell({
            type: new NumberCell(342, NumberFormat.INTEGER),
            style: new Style({ alignment: { horizontal: "right" }, width: 90 }),
          }),
          new Cell({
            type: new NumberCell(273600, NumberFormat.CURRENCY),
            style: new Style({
              alignment: { horizontal: "right" },
              width: 110,
            }),
          }),
          new Cell({
            type: new NumberCell(0.156, NumberFormat.PERCENTAGE),
            style: new Style({
              alignment: { horizontal: "right" },
              backgroundColor: "#d9ead3",
              width: 90,
            }),
          }),
          new Cell({
            type: new Checkbox(true),
            style: new Style({
              alignment: { horizontal: "center" },
              width: 80,
            }),
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Wireless Mouse") }),
          new Cell({
            type: new Dropdown({ values: statusOptions, selected: "Active" }),
          }),
          new Cell({
            type: new NumberCell(856, NumberFormat.INTEGER),
            style: new Style({ alignment: { horizontal: "right" } }),
          }),
          new Cell({
            type: new NumberCell(42800, NumberFormat.CURRENCY),
            style: new Style({ alignment: { horizontal: "right" } }),
          }),
          new Cell({
            type: new NumberCell(0.089, NumberFormat.PERCENTAGE),
            style: new Style({
              alignment: { horizontal: "right" },
              backgroundColor: "#fff2cc",
            }),
          }),
          new Cell({
            type: new Checkbox(true),
            style: new Style({ alignment: { horizontal: "center" } }),
          }),
        ],
      }),

      new HStack({
        children: [
          new Cell({ type: new Text("Mechanical Keyboard") }),
          new Cell({
            type: new Dropdown({ values: statusOptions, selected: "Pending" }),
          }),
          new Cell({
            type: new NumberCell(189, NumberFormat.INTEGER),
            style: new Style({ alignment: { horizontal: "right" } }),
          }),
          new Cell({
            type: new NumberCell(22680, NumberFormat.CURRENCY),
            style: new Style({ alignment: { horizontal: "right" } }),
          }),
          new Cell({
            type: new NumberCell(-0.023, NumberFormat.PERCENTAGE),
            style: new Style({
              alignment: { horizontal: "right" },
              backgroundColor: "#f4cccc",
            }),
          }),
          new Cell({
            type: new Checkbox(false),
            style: new Style({ alignment: { horizontal: "center" } }),
          }),
        ],
      }),
    ],
  });

  render(sheet, dashboard);
}
