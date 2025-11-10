# ReaSheets

ReaSheets is a declarative, component-based library for Google Apps Script that allows you to build complex table layouts in Google Sheets using a paradigm inspired by modern web frameworks like React.

## Motivation

Building complex, dynamic, and well-structured layouts directly with the Google Apps Script `SpreadsheetApp` API can be cumbersome. Managing cell positions, styling, merging, and data validation often leads to imperative and hard-to-maintain code. ReaSheets was built to solve this problem by providing a declarative API that lets you define your sheet layout as a tree of components, abstracting away the complexities of the underlying sheet manipulation.

## Core Concepts

The library is built around two main types of components:

1.  **Layout Components**: These components (`VStack`, `HStack`, `Cell`) define the structure and arrangement of your table. They can contain other components as children.
2.  **Type Components**: These components (`Text`, `Checkbox`, `Dropdown`, `DatePicker`) define the data type and behavior of a specific cell. They are used within `Cell` components.

You build a UI by composing these components into a tree, starting from a root layout component. This entire component tree is then passed to a `render` function, which translates your declarative layout into the appropriate Google Apps Script API calls.

## The Power of Composition: Creating Reusable Components

The true power of ReaSheets comes from composition. Because components are just simple JavaScript classes, you can create your own reusable components by wrapping them in functions. This allows you to build complex, domain-specific components from the simple primitives provided by the library.

For example, instead of defining the structure of a to-do list row repeatedly, you can create a `ToDoItem` component that encapsulates all the logic and styling for a single row:

```javascript
function ToDoItem({ id, description, status, isDone, isEven = false }) {
    const rowStyle = isEven ? new Style({ backgroundColor: '#f3f3f3' }) : new Style();

    return new HStack({
        style: rowStyle,
        children: [
            new Cell({ type: new Text(id.split('-')[0]), style: new Style({ font: { bold: true } }) }),
            new Cell({ type: new Text(id.split('-')[1]) }),
            new Cell({ type: new Text(description), colSpan: 3 }),
            new Cell({
                type: new Dropdown({
                    values: ['Pending', 'In Progress', 'Complete'],
                    selected: status
                })
            }),
            new Cell({ type: new DatePicker({ format: 'yyyy-mm-dd' }) }),
            new Cell({ type: new Checkbox(isDone) })
        ]
    });
}
```

Now, your main layout becomes much cleaner and easier to read, as shown in the **Usage Example** below. This pattern is highly encouraged and is key to managing the complexity of large tables.

## The Render Pipeline

The rendering process is handled by two specialized engines that work in sequence.

### 1. The Layout Engine

The Layout Engine's job is to translate your hierarchical component tree into a flat array of "resolved cells," where each cell has an absolute row and column position.

It works by recursively walking the component tree, starting from the root:

1.  Each layout component (`VStack`, `HStack`) receives a starting position (`{row, col}`) and calculates the positions of its children.
2.  A crucial part of this process is the **`occupancyMap`**, a `Set` that keeps track of every cell coordinate that is already in use, including those covered by row or column spans.
3.  When placing a child component, the engine checks the `occupancyMap`. If the target cell is already occupied, it advances the cursor to the next available cell (either to the next row in a `VStack` or the next column in an `HStack`). This is the primary mechanism for resolving cell spanning conflicts and ensuring that components do not overlap.

### 2. The Commit Engine

Once the Layout Engine produces a flat array of resolved cells, the Commit Engine takes over. Its goal is to apply this layout to the actual Google Sheet as efficiently as possible.

1.  **Grid Creation**: The engine first determines the bounding box of the entire layout. It then creates 2D arrays (grids) for every style property (backgrounds, font colors, font sizes, etc.).
2.  **Grid Population**: It iterates through the resolved cells and populates these grids with the final, calculated values for each cell. During this step, it calls the `getRenderDirectives()` method on each cell's `Type` component to get any special validation or formatting rules.
3.  **Batched API Calls**: Finally, the engine makes a series of highly efficient, batched API calls to `SpreadsheetApp`, setting all values, backgrounds, font styles, etc., for the entire range at once. This is significantly faster than setting properties cell by cell. Merges and borders are handled in separate passes.

### Conflict and Style Resolution

-   **Cell Spanning**: As described above, the `occupancyMap` in the Layout Engine prevents conflicts. A component will never be rendered on top of a cell that is already occupied by another component or is part of another component's span.
-   **Style Merging**: Styles are inherited from parent to child. If a `VStack` has a font style, all children inside it will inherit that style. A child can override any inherited style by providing its own `Style` prop. The child's style properties always take precedence.
-   **Border Merging**: Borders are also inherited. If a parent `HStack` defines a bottom border, and a `Cell` within it defines its own bottom border, the `Cell`'s border will override the parent's.

## Component API

### Layout Components

#### `VStack`

Arranges its children vertically in a single column.

-   **`children`**: `Array<Component>` (Required) - An array of `HStack`, `VStack`, or `Cell` components.
-   **`style`**: `Style` (Optional) - A style to apply to all children within the stack.

#### `HStack`

Arranges its children horizontally in a single row.

-   **`children`**: `Array<Component>` (Required) - An array of `HStack`, `VStack`, or `Cell` components.
-   **`style`**: `Style` (Optional) - A style to apply to all children within the stack.

#### `Cell`

Represents a single cell or a merged block of cells in the sheet.

-   **`type`**: `Type` (Optional) - The data type of the cell (e.g., `new Text('Hello')`). Defaults to an empty `Text` component.
-   **`style`**: `Style` (Optional) - A style unique to this cell.
-   **`note`**: `string` (Optional) - A note to attach to the cell.
-   **`colSpan`**: `number` (Optional) - The number of columns the cell should span. Defaults to `1`.
-   **`rowSpan`**: `number` (Optional) - The number of rows the cell should span. Defaults to `1`.

### Type Components

#### `Text`

A simple text value.

-   **`value`**: `string` (Optional) - The text to display.

#### `Checkbox`

A checkbox.

-   **`isChecked`**: `boolean` (Optional) - The initial checked state. Defaults to `false`.

#### `DatePicker`

A cell with date validation.

-   **`format`**: `string` (Optional) - A number format string for the date (e.g., `'yyyy-mm-dd'`).

#### `Dropdown`

A cell with a dropdown list.

-   **`values`**: `Array<string | object>` (Required) - An array of options.
    -   If an array of strings, it creates a simple dropdown.
    -   If an array of objects `{ value: string, style: Style }`, it creates a dropdown where each option has its own conditional formatting rule.
-   **`selected`**: `string` (Optional) - The initially selected value. Defaults to the first item in the `values` array.

### Styling

#### `Style`

Defines the visual style of a component. All properties are optional.

-   **`backgroundColor`**: `string` - A CSS color string (e.g., `'#ff0000'`, `'blue'`).
-   **`font`**: `object` - An object with font properties:
    -   `color`: `string`
    -   `size`: `number`
    -   `family`: `string`
    -   `bold`: `boolean`
    -   `italic`: `boolean`
    -   `underline`: `boolean`
    -   `strikethrough`: `boolean`
-   **`alignment`**: `object` - An object with alignment properties:
    -   `horizontal`: `string` (e.g., `'left'`, `'center'`, `'right'`)
    -   `vertical`: `string` (e.g., `'top'`, `'middle'`, `'bottom'`)
-   **`wrap`**: `object` - An object with a `strategy` property from the `WrapStrategy` enum.
-   **`border`**: `Border` - A `Border` object.
-   **`rotation`**: `object` - An object with an `angle` property (`-90` to `90`).

#### `Border`

Defines the borders for a `Style`.

-   **`top` / `bottom` / `left` / `right`**: `object` (Optional) - An object defining a border side:
    -   `color`: `string` (Required)
    -   `thickness`: `BorderThickness` (Required) - A value from the `BorderThickness` enum.

#### Enums

-   **`BorderThickness`**: Maps to `SpreadsheetApp.BorderStyle`.
    -   `DOTTED`, `DASHED`, `SOLID`, `SOLID_MEDIUM`, `SOLID_THICK`, `DOUBLE`
-   **`WrapStrategy`**: Maps to `SpreadsheetApp.WrapStrategy`.
    -   `WRAP`, `OVERFLOW`, `CLIP`

## Usage Example

Here is a complete example demonstrating how to define a layout, create a reusable component, and render a to-do list.

```javascript
// 1. Define a reusable component for a single to-do list item
function ToDoItem({ id, description, status, isDone, isEven = false }) {
    const rowStyle = isEven ? new Style({ backgroundColor: '#f3f3f3' }) : new Style();

    // Dropdown options with conditional formatting
    const statusOptions = [
        { value: 'Pending', style: new Style({ backgroundColor: '#fff2cc' }) },
        { value: 'In Progress', style: new Style({ backgroundColor: '#cfe2f3' }) },
        { value: 'Complete', style: new Style({ backgroundColor: '#d9ead3' }) }
    ];

    return new HStack({
        style: rowStyle,
        children: [
            new Cell({ type: new Text(id.split('-')[0]), style: new Style({ font: { bold: true } }) }),
            new Cell({ type: new Text(id.split('-')[1]) }),
            new Cell({ type: new Text(description), colSpan: 3 }),
            new Cell({
                type: new Dropdown({
                    values: statusOptions,
                    selected: status
                })
            }),
            new Cell({ type: new DatePicker({ format: 'yyyy-mm-dd' }) }),
            new Cell({ type: new Checkbox(isDone) })
        ]
    });
}

/**
 * The main function to generate the sheet.
 */
function createToDoListSheet() {
    // 2. Get the target sheet and clear it
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ToDoList') || ss.insertSheet('ToDoList');
    sheet.clear();
    sheet.setFrozenRows(1);

    // 3. Define a style for the header
    const headerStyle = new Style({
        backgroundColor: '#4a86e8',
        font: { color: 'white', bold: true },
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: new Border({
            bottom: { color: 'black', thickness: BorderThickness.SOLID_THICK }
        })
    });
    
    const header = new HStack({
        style: headerStyle,
        children: [
            new Cell({ type: new Text('TASK ID'), colSpan: 2 }),
            new Cell({ type: new Text('DESCRIPTION'), colSpan: 3 }),
            new Cell({ type: new Text('STATUS') }),
            new Cell({ type: new Text('DUE DATE') }),
            new Cell({ type: new Text('DONE') })
        ]
    });

    // 4. Define the main layout using the reusable ToDoItem component
    const myToDoList = new VStack({
        style: new Style({ font: { family: 'Arial', size: 10 } }),
        children: [
            header,
            ToDoItem({ id: 'PROJ-A-101', description: 'Finalize Q3 report.', status: 'In Progress', isDone: false }),
            ToDoItem({ id: 'PROJ-B-205', description: 'Onboard new team members.', status: 'Pending', isDone: false, isEven: true }),
            ToDoItem({ id: 'PROJ-A-102', description: 'Prepare slides for stakeholder meeting.', status: 'Complete', isDone: true }),
        ]
    });

    // 5. Render the final layout to the sheet
    render(myToDoList, sheet);
    
    // Adjust column widths for better readability
    sheet.autoResizeColumns(1, sheet.getLastColumn());
}
```
