# ReaSheet

ReaSheets is a declarative, component-based library for Google Apps Script that brings the power of component composition (like React) to Google Sheets.

**Design Philosophy:**
*   **Single File:** Intentionally designed as a single file (`ReaSheet.js`) so you can copy-paste it into any Apps Script project without build tools.
*   **Declarative:** Describe *what* your sheet should look like, not *how* to build it cell-by-cell.
*   **Composable:** Build complex UIs from simple, reusable components.
*   **Performance:** Uses a virtual layout engine to minimize slow Google Apps Script API calls.

## How it Functions

ReaSheet separates the definition of your sheet from the execution of the update. It runs in two distinct phases:

### 1. The Layout Phase (Virtual Calculation)
When you run `render(layout, sheet)`, the library first builds a "Virtual Sheet" in memory.
*   **Recursive Layout:** It walks down your component tree (`VStack` -> `HStack` -> `Cell`).
*   **Collision Detection:** It maintains an `OccupancyMap` of every used cell. If you try to place a cell in a spot taken by a previous `rowSpan` or `colSpan`, the cursor automatically advances to the next available slot.
*   **Result:** This phase produces a flat list of "Resolved Cells" with absolute coordinates (e.g., "Row 5, Col 3: Text 'Hello', Red Background").

### 2. The Commit Phase (Batched Execution)
Once the layout is calculated, the `Renderer` translates these resolved cells into the minimum number of Google Apps Script API calls.
*   **Grid Building:** It generates 2D arrays for every property (Values, Backgrounds, FontColors, Validations, etc.).
*   **Batching:** Instead of calling `cell.setValue()` 1000 times, it calls `range.setValues(grid)` once. This effectively solves the "Time Limit Exceeded" errors common in complex scripts.

## Basic Usage

The library exposes its components globally, so you can just use them directly.

```javascript
function createSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Demo") || ss.insertSheet("Demo");
  sheet.clear();

  // 1. Define Styles
  const headerStyle = new Style({
    backgroundColor: "#4a86e8",
    font: { color: "white", bold: true },
    alignment: { horizontal: "center" }
  });

  // 2. Build Layout
  const layout = new VStack({
    children: [
      // Header Row
      new HStack({
        style: headerStyle,
        children: [
          new Cell({ type: new Text("Item"), width: 200 }),
          new Cell({ type: new Text("Status"), width: 150 }),
        ],
      }),
      
      // Data Row
      new HStack({
        children: [
          new Cell({ type: new Text("Project Alpha") }),
          new Cell({ type: new Text("Active") }),
        ],
      })
    ]
  });

  // 3. Render
  render(layout, sheet);
}
```

## Component API

### Layout Components
*   **`VStack({ children, style })`**: Stacks components vertically. It fills rows top-to-bottom.
*   **`HStack({ children, style })`**: Stacks components horizontally. It fills columns left-to-right.
*   **`Cell({ type, style, rowSpan, colSpan })`**: The atomic unit. Handles spanning and content.

### Data Types
These are passed to the `type` prop of a `Cell`.
*   **`Text(value)`**: Simple text string.
*   **`NumberCell(value, format)`**: Numeric value with format pattern (e.g., `NumberFormats.CURRENCY`).
*   **`Checkbox(checked)`**: Boolean checkbox validation.
*   **`Dropdown({ values, selected })`**: Data validation dropdown. `values` can be simple strings or objects with conditional formatting styles.
*   **`DatePicker({ format })`**: Date validation and formatting.

### Styling
*   **`Style({ ... })`**: The styling object. Properties:
    *   `backgroundColor`: Hex code.
    *   `font`: `{ color, size, bold, italic, ... }`
    *   `alignment`: `{ horizontal, vertical }`
    *   `border`: `new Border({ top: { color, thickness }, ... })`
    *   `wrap`: `WrapStrategy.WRAP` | `OVERFLOW` | `CLIP`
*   **Inheritance:** Styles cascade down. A `Style` on a `VStack` applies to all its children unless overridden.
