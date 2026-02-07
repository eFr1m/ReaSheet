/**
 * ReaSheets - Declarative UI for Google Sheets
 *
 * A lightweight, component-based library for Google Apps Script.
 */

// ============================================================================
// ENUMS
// ============================================================================

var WrapStrategy = Object.freeze({
  WRAP: SpreadsheetApp.WrapStrategy.WRAP,
  OVERFLOW: SpreadsheetApp.WrapStrategy.OVERFLOW,
  CLIP: SpreadsheetApp.WrapStrategy.CLIP,
});

var BorderStyle = Object.freeze({
  DOTTED: SpreadsheetApp.BorderStyle.DOTTED,
  DASHED: SpreadsheetApp.BorderStyle.DASHED,
  SOLID: SpreadsheetApp.BorderStyle.SOLID,
  SOLID_MEDIUM: SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
  SOLID_THICK: SpreadsheetApp.BorderStyle.SOLID_THICK,
  DOUBLE: SpreadsheetApp.BorderStyle.DOUBLE,
});

var HAlign = Object.freeze({
  LEFT: "left",
  CENTER: "center",
  RIGHT: "right",
});

var VAlign = Object.freeze({
  TOP: "top",
  MIDDLE: "middle",
  BOTTOM: "bottom",
});

var NumberFormat = Object.freeze({
  PERCENTAGE: "0.00%",
  CURRENCY: "$#,##0.00",
  INTEGER: "0",
  DECIMAL: "0.00",
  DATE: "dd/MM/yyyy",
});

// ============================================================================
// BORDER
// ============================================================================

class Border {
  constructor({ top = null, bottom = null, left = null, right = null } = {}) {
    this.top = top;
    this.bottom = bottom;
    this.left = left;
    this.right = right;
    Object.freeze(this);
  }

  static all(color, style = BorderStyle.SOLID) {
    const side = { color, style };
    return new Border({ top: side, bottom: side, left: side, right: side });
  }

  static none() {
    return new Border();
  }

  equals(other) {
    if (this === other) return true;
    if (!other) return false;
    return (
      this._sideEquals(this.top, other.top) &&
      this._sideEquals(this.bottom, other.bottom) &&
      this._sideEquals(this.left, other.left) &&
      this._sideEquals(this.right, other.right)
    );
  }

  _sideEquals(a, b) {
    if (a === b) return true;
    if (!a || !b) return false;
    return a.color === b.color && a.style === b.style;
  }
}

// ============================================================================
// STYLE
// ============================================================================

var _defaultStyle = {
  backgroundColor: null,
  font: {
    color: "black",
    size: 10,
    family: "Arial",
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
  },
  alignment: {
    horizontal: HAlign.LEFT,
    vertical: VAlign.TOP,
  },
  wrap: WrapStrategy.OVERFLOW,
  border: new Border(),
  rotation: 0,
  width: null,
  height: null,
};

class Style {
  constructor({
    backgroundColor = null,
    font = {},
    alignment = {},
    wrap = WrapStrategy.OVERFLOW,
    border = new Border(),
    rotation = 0,
    width = null,
    height = null,
  } = {}) {
    this.backgroundColor = backgroundColor;
    this.font = { ..._defaultStyle.font, ...font };
    this.alignment = { ..._defaultStyle.alignment, ...alignment };
    this.wrap = wrap;
    this.border = border;
    this.rotation = rotation;
    this.width = width;
    this.height = height;
    Object.freeze(this);
  }

  merge(child) {
    if (!child) return this;
    return new Style({
      backgroundColor: child.backgroundColor ?? this.backgroundColor,
      font: { ...this.font, ...child.font },
      alignment: { ...this.alignment, ...child.alignment },
      wrap: child.wrap ?? this.wrap,
      border: child.border ?? this.border,
      rotation: child.rotation ?? this.rotation,
      width: child.width ?? this.width,
      height: child.height ?? this.height,
    });
  }
}

// ============================================================================
// DATA TYPES (Cell Content)
// ============================================================================

class Text {
  constructor(value = "") {
    this.value = value;
  }

  getDirectives() {
    return {};
  }
}

class NumberCell {
  constructor(value, format = "0") {
    this.value = value;
    this.format = format;
  }

  getDirectives() {
    return { numberFormat: this.format };
  }
}

class Checkbox {
  constructor(checked = false) {
    this.value = checked;
  }

  getDirectives() {
    return {
      validation: SpreadsheetApp.newDataValidation().requireCheckbox().build(),
    };
  }
}

class Dropdown {
  constructor({ values, selected = null }) {
    const isObjectArray = values[0]?.value !== undefined;

    this.values = values;
    this.plainValues = isObjectArray ? values.map((v) => v.value) : values;
    this.value = selected ?? this.plainValues[0];
    this.isObjectArray = isObjectArray;
  }

  getDirectives(range) {
    const directives = {
      validation: SpreadsheetApp.newDataValidation()
        .requireValueInList(this.plainValues)
        .build(),
    };

    if (this.isObjectArray) {
      directives.conditionalFormatRules = this.values
        .filter((item) => item.style)
        .map((item) =>
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(item.value)
            .setBackground(item.style.backgroundColor)
            .setFontColor(item.style.font?.color)
            .setRanges([range])
            .build(),
        );
    }
    return directives;
  }
}

class DatePicker {
  constructor(arg = {}) {
    if (arg instanceof Date) {
      this.format = NumberFormat.DATE;
      this.value = arg;
      return;
    }

    const { format = NumberFormat.DATE, value = null } = arg || {};
    if (value !== null && !(value instanceof Date)) {
      throw new TypeError("DatePicker value must be a Date or null");
    }

    this.format = format;
    this.value = value;
  }

  getDirectives() {
    return {
      validation: SpreadsheetApp.newDataValidation().requireDate().build(),
      numberFormat: this.format,
    };
  }

  get serialValue() {
    if (!this.value) return "";
    return this.value;
  }
}

// ============================================================================
// COMPONENTS
// ============================================================================

class Cell {
  constructor({
    type = new Text(""),
    style = null,
    note = "",
    colSpan = 1,
    rowSpan = 1,
  }) {
    this.type = type;
    this.style = style;
    this.note = note;
    this.colSpan = colSpan;
    this.rowSpan = rowSpan;
  }

  render(ctx, pos, inheritedStyle) {
    const finalStyle = inheritedStyle.merge(this.style);

    // Mark occupied cells
    for (let r = 0; r < this.rowSpan; r++) {
      for (let c = 0; c < this.colSpan; c++) {
        ctx.occupied.add(`${pos.row + r}:${pos.col + c}`);
      }
    }

    return [
      {
        row: pos.row,
        col: pos.col,
        cell: this,
        style: finalStyle,
      },
    ];
  }
}

class HStack {
  constructor({ children, style = null }) {
    this.children = children;
    this.style = style;
  }

  render(ctx, pos, inheritedStyle) {
    const containerStyle = inheritedStyle.merge(this.style);
    const resolved = [];
    let col = pos.col;

    for (const child of this.children) {
      // Skip occupied cells
      while (ctx.occupied.has(`${pos.row}:${col}`)) col++;

      const childCells = child.render(
        ctx,
        { row: pos.row, col },
        containerStyle,
      );
      resolved.push(...childCells);

      // Advance past this child
      let maxCol = col;
      for (const c of childCells) {
        maxCol = Math.max(maxCol, c.col + (c.cell.colSpan || 1) - 1);
      }
      col = maxCol + 1;
    }

    return resolved;
  }
}

class VStack {
  constructor({ children, style = null }) {
    this.children = children;
    this.style = style;
  }

  render(ctx, pos, inheritedStyle) {
    const containerStyle = inheritedStyle.merge(this.style);
    const resolved = [];
    let row = pos.row;

    for (const child of this.children) {
      // Skip occupied rows
      while (ctx.occupied.has(`${row}:${pos.col}`)) row++;

      const childCells = child.render(
        ctx,
        { row, col: pos.col },
        containerStyle,
      );
      resolved.push(...childCells);

      // Advance past this child
      let maxRow = row;
      for (const c of childCells) {
        maxRow = Math.max(maxRow, c.row + (c.cell.rowSpan || 1) - 1);
      }
      row = maxRow + 1;
    }

    return resolved;
  }
}
// ============================================================================
// RENDERER (Sheets API v4 - Strict Mode)
// ============================================================================

function render(sheet, root) {
  const isSheet = (value) =>
    value &&
    typeof value.getSheetId === "function" &&
    typeof value.getParent === "function";

  const isRoot = (value) => value && typeof value.render === "function";

  if (!isSheet(sheet) || !isRoot(root)) {
    throw new TypeError("render expects (sheet, root)");
  }

  // 1. Calculate Layout
  const ctx = { occupied: new Set() };
  // API uses 0-based indexing.
  const cells = root.render(ctx, { row: 0, col: 0 }, new Style());

  if (cells.length === 0) return;

  const bounds = _calculateBounds(cells);
  const sheetId = sheet.getSheetId();
  const spreadsheetId = sheet.getParent().getId();

  // 2. Flush any pending SpreadsheetApp operations (e.g. sheet.clear())
  // to prevent them from executing after our batchUpdate and wiping data.
  SpreadsheetApp.flush();

  // 3. Build API Payloads
  const requests = [];

  // A. Unmerge any existing merges in the target range
  requests.push({
    unmergeCells: {
      range: {
        sheetId: sheetId,
        startRowIndex: bounds.minRow,
        endRowIndex: bounds.minRow + bounds.numRows,
        startColumnIndex: bounds.minCol,
        endColumnIndex: bounds.minCol + bounds.numCols,
      },
    },
  });

  // B. Clear the range to remove artifacts
  requests.push({
    updateCells: {
      range: {
        sheetId: sheetId,
        startRowIndex: bounds.minRow,
        endRowIndex: bounds.minRow + bounds.numRows,
        startColumnIndex: bounds.minCol,
        endColumnIndex: bounds.minCol + bounds.numCols,
      },
      fields: "userEnteredValue,userEnteredFormat,note,dataValidation",
    },
  });

  // C. Generate Cell Data
  const { rows, conditionalRules } = _buildApiData(sheet, cells, bounds);

  // D. Update Content & Formatting
  requests.push({
    updateCells: {
      start: {
        sheetId: sheetId,
        rowIndex: bounds.minRow,
        columnIndex: bounds.minCol,
      },
      rows: rows,
      // STRICT field mask ensures we apply exactly what we generated
      fields: "userEnteredValue,userEnteredFormat,note,dataValidation",
    },
  });

  // E. Merges
  const merges = _buildApiMerges(cells, bounds, sheetId);
  if (merges.length) requests.push(...merges);

  // F. Dimensions (Widths/Heights)
  const dims = _buildApiDimensions(cells, bounds, sheetId);
  if (dims.length) requests.push(...dims);

  // 4. Send Batch Request
  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
  }

  // 5. Apply Conditional Rules (via AppScript)
  if (conditionalRules.length > 0) {
    const existing = sheet.getConditionalFormatRules();
    sheet.setConditionalFormatRules(existing.concat(conditionalRules));
  }
}

// ----------------------------------------------------------------------------
// DATA BUILDERS
// ----------------------------------------------------------------------------

function _calculateBounds(cells) {
  let minRow = Infinity,
    maxRow = 0,
    minCol = Infinity,
    maxCol = 0;
  for (const c of cells) {
    minRow = Math.min(minRow, c.row);
    maxRow = Math.max(maxRow, c.row + (c.cell.rowSpan || 1) - 1);
    minCol = Math.min(minCol, c.col);
    maxCol = Math.max(maxCol, c.col + (c.cell.colSpan || 1) - 1);
  }
  return {
    minRow,
    maxRow,
    minCol,
    maxCol,
    numRows: maxRow - minRow + 1,
    numCols: maxCol - minCol + 1,
  };
}

function _buildApiData(sheet, cells, bounds) {
  const { minRow, minCol, numRows, numCols } = bounds;

  // Initialize a dense grid of empty objects
  const grid = Array.from({ length: numRows }, () =>
    Array.from({ length: numCols }, () => ({})),
  );

  const conditionalRules = [];

  for (const c of cells) {
    const r = c.row - minRow;
    const col = c.col - minCol;

    // Directives (Validation, NumberFormat)
    let directives = {};
    if (c.cell.type.getDirectives) {
      // Convert 0-based to 1-based for getRange if needed by directives
      const rng = sheet.getRange(
        c.row + 1,
        c.col + 1,
        c.cell.rowSpan || 1,
        c.cell.colSpan || 1,
      );
      directives = c.cell.type.getDirectives(rng);
    }

    if (directives.conditionalFormatRules) {
      conditionalRules.push(...directives.conditionalFormatRules);
    }

    // Only populate the top-left cell of the merge
    // (We check if the slot is empty to avoid overwriting if logic is flawed,
    // though the render tree should guarantee uniqueness)
    if (Object.keys(grid[r][col]).length === 0) {
      grid[r][col] = _createCellData(c, directives);
    }
  }

  // Convert grid to API RowData structure
  const rows = grid.map((rowValues) => ({ values: rowValues }));

  return { rows, conditionalRules };
}

function _createCellData({ cell, style }, directives) {
  const cellData = {};

  // 1. Value
  const v = cell.type.value;
  // Explicitly check against null/undefined. Empty string "" is a valid value.
  if (v !== null && v !== undefined) {
    if (typeof v === "number") cellData.userEnteredValue = { numberValue: v };
    else if (typeof v === "boolean")
      cellData.userEnteredValue = { boolValue: v };
    else if (v instanceof Date)
      cellData.userEnteredValue = { numberValue: _dateToSerial(v) };
    else cellData.userEnteredValue = { stringValue: String(v) };
  }

  // 2. Format
  // We MUST attach userEnteredFormat even if empty, otherwise the API might not clear old formats correctly
  const format = _mapStyleToFormat(style);

  // Apply Number Format from Directives
  if (directives.numberFormat) {
    const isDate =
      directives.numberFormat.includes("d") ||
      directives.numberFormat.includes("y");
    format.numberFormat = {
      type: isDate ? "DATE" : "NUMBER",
      pattern: directives.numberFormat,
    };
  }

  cellData.userEnteredFormat = format;

  // 3. Validation
  if (directives.validation) {
    cellData.dataValidation = _mapValidationToApi(directives.validation);
  }

  // 4. Note
  if (cell.note) cellData.note = cell.note;

  return cellData;
}

function _mapStyleToFormat(style) {
  const format = {};

  // Background
  const bg = _parseColor(style.backgroundColor);
  if (bg) format.backgroundColor = bg;

  // Alignment
  if (style.alignment.horizontal)
    format.horizontalAlignment = style.alignment.horizontal.toUpperCase();
  if (style.alignment.vertical)
    format.verticalAlignment = style.alignment.vertical.toUpperCase();

  // Wrap Strategy
  const wrapMap = {
    [SpreadsheetApp.WrapStrategy.WRAP]: "WRAP",
    [SpreadsheetApp.WrapStrategy.OVERFLOW]: "OVERFLOW_CELL",
    [SpreadsheetApp.WrapStrategy.CLIP]: "CLIP",
  };
  // Default to OVERFLOW_CELL if the mapping fails or style is missing
  format.wrapStrategy = wrapMap[style.wrap] || "OVERFLOW_CELL";

  // Text Format
  const tf = {};
  const fg = _parseColor(style.font.color);
  if (fg) tf.foregroundColor = fg;

  if (style.font.family) tf.fontFamily = style.font.family;
  if (style.font.size) tf.fontSize = style.font.size;
  if (style.font.bold) tf.bold = true;
  if (style.font.italic) tf.italic = true;
  if (style.font.strikethrough) tf.strikethrough = true;
  if (style.font.underline) tf.underline = true;

  // Only attach textFormat if we have properties
  if (Object.keys(tf).length > 0) format.textFormat = tf;

  // Borders
  const borders = {};
  if (style.border.top) borders.top = _mapBorder(style.border.top);
  if (style.border.bottom) borders.bottom = _mapBorder(style.border.bottom);
  if (style.border.left) borders.left = _mapBorder(style.border.left);
  if (style.border.right) borders.right = _mapBorder(style.border.right);

  if (Object.keys(borders).length > 0) format.borders = borders;

  // Rotation
  if (style.rotation) format.textRotation = { angle: style.rotation };

  return format;
}

function _mapBorder(side) {
  if (!side) return null;
  const b = {
    style: side.style,
    color: _parseColor(side.color),
  };
  // API border objects cannot have partial nulls
  if (!b.color) delete b.color;
  return b;
}

function _mapValidationToApi(validation) {
  const criteria = validation.getCriteriaType();
  const args = validation.getCriteriaValues();

  const rule = {
    showCustomUi: true,
    strict: !validation.getAllowInvalid(),
  };

  if (criteria === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
    rule.condition = { type: "BOOLEAN" };
  } else if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    rule.condition = {
      type: "ONE_OF_LIST",
      values: args[0].map((val) => ({ userEnteredValue: String(val) })),
    };
  } else if (
    criteria === SpreadsheetApp.DataValidationCriteria.DATE_IS_VALID_DATE
  ) {
    rule.condition = { type: "DATE_IS_VALID" };
  }

  return rule;
}

function _buildApiMerges(cells, bounds, sheetId) {
  const requests = [];
  for (const c of cells) {
    if (c.cell.rowSpan > 1 || c.cell.colSpan > 1) {
      requests.push({
        mergeCells: {
          range: {
            sheetId: sheetId,
            startRowIndex: c.row,
            endRowIndex: c.row + c.cell.rowSpan,
            startColumnIndex: c.col,
            endColumnIndex: c.col + c.cell.colSpan,
          },
          mergeType: "MERGE_ALL",
        },
      });
    }
  }
  return requests;
}

function _buildApiDimensions(cells, bounds, sheetId) {
  const requests = [];
  const widthMap = new Map();
  const heightMap = new Map();

  for (const c of cells) {
    if (c.style.width !== null) widthMap.set(c.col, c.style.width);
    if (c.style.height !== null) heightMap.set(c.row, c.style.height);
  }

  widthMap.forEach((width, index) => {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId,
          dimension: "COLUMNS",
          startIndex: index,
          endIndex: index + 1,
        },
        properties: { pixelSize: width },
        fields: "pixelSize",
      },
    });
  });

  heightMap.forEach((height, index) => {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId,
          dimension: "ROWS",
          startIndex: index,
          endIndex: index + 1,
        },
        properties: { pixelSize: height },
        fields: "pixelSize",
      },
    });
  });

  return requests;
}

// ----------------------------------------------------------------------------
// UTILS
// ----------------------------------------------------------------------------

function _dateToSerial(date) {
  const epoch = new Date(1899, 11, 30);
  const msPerDay = 86400000;
  return (
    (date.getTime() - epoch.getTime() - date.getTimezoneOffset() * 60000) /
    msPerDay
  );
}

function _parseColor(input) {
  if (!input) return null;

  const hexMap = {
    black: "#000000",
    white: "#FFFFFF",
    red: "#FF0000",
    blue: "#0000FF",
    green: "#008000",
    gray: "#808080",
    grey: "#808080",
    yellow: "#FFFF00",
    orange: "#FFA500",
    purple: "#800080",
  };

  let hex = hexMap[input.toLowerCase()] || input;

  if (typeof hex === "string" && hex.startsWith("#")) {
    // Expand shorthand hex (e.g. #fff -> #ffffff)
    if (hex.length === 4) {
      hex = "#" + hex[1] + hex[1] + hex[2] + hex[2] + hex[3] + hex[3];
    }

    if (hex.length === 7) {
      const r = parseInt(hex.slice(1, 3), 16) / 255;
      const g = parseInt(hex.slice(3, 5), 16) / 255;
      const b = parseInt(hex.slice(5, 7), 16) / 255;

      // Ensure valid numbers
      if (!isNaN(r) && !isNaN(g) && !isNaN(b)) {
        return { red: r, green: g, blue: b };
      }
    }
  }
  return null;
}
