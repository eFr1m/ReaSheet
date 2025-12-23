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
            .build()
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
        containerStyle
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
        containerStyle
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
// RENDERER
// ============================================================================

function render(sheet, root) {
  const ctx = { occupied: new Set() };
  const cells = root.render(ctx, { row: 1, col: 1 }, new Style());

  if (cells.length === 0) return;

  const bounds = _calculateBounds(cells);
  const range = sheet.getRange(
    bounds.minRow,
    bounds.minCol,
    bounds.numRows,
    bounds.numCols
  );
  range.clear();

  const grids = _buildGrids(cells, bounds);

  // Bulk apply styles
  range
    .setValues(grids.values)
    .setNotes(grids.notes)
    .setBackgrounds(grids.backgrounds)
    .setFontColors(grids.fontColors)
    .setFontSizes(grids.fontSizes)
    .setFontWeights(grids.fontWeights)
    .setFontStyles(grids.fontStyles)
    .setFontLines(grids.fontLines)
    .setHorizontalAlignments(grids.hAligns)
    .setVerticalAlignments(grids.vAligns)
    .setWrapStrategies(grids.wraps)
    .setNumberFormats(grids.numberFormats)
    .setDataValidations(grids.validations);

  // Rotations (only if needed)
  if (grids.hasRotation) {
    range.setTextRotations(grids.rotations);
  }

  // Dimensions
  for (const [col, width] of Object.entries(grids.widths)) {
    sheet.setColumnWidth(parseInt(col), width);
  }
  for (const [row, height] of Object.entries(grids.heights)) {
    sheet.setRowHeight(parseInt(row), height);
  }

  // Borders (RLE optimized)
  _applyBorders(sheet, bounds, grids.borders);

  // Merges
  for (const m of grids.merges) {
    sheet.getRange(m.row, m.col, m.rowSpan, m.colSpan).merge();
  }

  // Conditional formats (merge with existing)
  if (grids.conditionalRules.length > 0) {
    const existing = sheet.getConditionalFormatRules();
    sheet.setConditionalFormatRules(existing.concat(grids.conditionalRules));
  }
}

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

function _buildGrids(cells, bounds) {
  const { minRow, minCol, numRows, numCols } = bounds;
  const grid = (fill) =>
    Array.from({ length: numRows }, () => Array(numCols).fill(fill));

  const grids = {
    values: grid(""),
    notes: grid(""),
    backgrounds: grid(null),
    fontColors: grid(null),
    fontSizes: grid(null),
    fontWeights: grid(null),
    fontStyles: grid(null),
    fontLines: grid(null),
    hAligns: grid(null),
    vAligns: grid(null),
    wraps: grid(WrapStrategy.OVERFLOW),
    numberFormats: grid("General"),
    validations: grid(null),
    rotations: grid(0),
    borders: grid(null),
    widths: {},
    heights: {},
    merges: [],
    conditionalRules: [],
    hasRotation: false,
  };

  for (const c of cells) {
    const { row, col, cell, style } = c;
    const { type, note, rowSpan, colSpan } = cell;
    const directives =
      type.getDirectives?.(
        SpreadsheetApp.getActiveSpreadsheet()
          .getActiveSheet()
          .getRange(row, col, rowSpan, colSpan)
      ) || {};

    // Track dimensions
    if (style.width !== null) grids.widths[col] = style.width;
    if (style.height !== null) grids.heights[row] = style.height;

    // Conditional rules
    if (directives.conditionalFormatRules) {
      grids.conditionalRules.push(...directives.conditionalFormatRules);
    }

    // Merges
    if (rowSpan > 1 || colSpan > 1) {
      grids.merges.push({ row, col, rowSpan, colSpan });
    }

    // Fill grid cells
    for (let rOff = 0; rOff < rowSpan; rOff++) {
      for (let cOff = 0; cOff < colSpan; cOff++) {
        const r = row - minRow + rOff;
        const c_idx = col - minCol + cOff;
        const isTopLeft = rOff === 0 && cOff === 0;

        grids.values[r][c_idx] = isTopLeft ? (type.serialValue ?? type.value) : "";
        grids.notes[r][c_idx] = isTopLeft ? note : "";
        grids.backgrounds[r][c_idx] = style.backgroundColor;
        grids.fontColors[r][c_idx] = style.font.color;
        grids.fontSizes[r][c_idx] = style.font.size;
        grids.fontWeights[r][c_idx] = style.font.bold ? "bold" : "normal";
        grids.fontStyles[r][c_idx] = style.font.italic ? "italic" : "normal";
        grids.fontLines[r][c_idx] = style.font.underline
          ? "underline"
          : style.font.strikethrough
          ? "line-through"
          : null;
        grids.hAligns[r][c_idx] = style.alignment.horizontal;
        grids.vAligns[r][c_idx] = style.alignment.vertical;
        grids.wraps[r][c_idx] = style.wrap;
        grids.rotations[r][c_idx] = style.rotation;
        grids.validations[r][c_idx] = directives.validation || null;

        if (directives.numberFormat) {
          grids.numberFormats[r][c_idx] = directives.numberFormat;
        }

        if (style.rotation !== 0) grids.hasRotation = true;

        // Border: only apply outer edges
        const border = style.border;
        grids.borders[r][c_idx] = {
          top: rOff === 0 ? border.top : null,
          bottom: rOff === rowSpan - 1 ? border.bottom : null,
          left: cOff === 0 ? border.left : null,
          right: cOff === colSpan - 1 ? border.right : null,
        };
      }
    }
  }

  return grids;
}

function _applyBorders(sheet, bounds, borders) {
  const { minRow, minCol, numRows, numCols } = bounds;

  for (let r = 0; r < numRows; r++) {
    let c = 0;
    while (c < numCols) {
      const b = borders[r][c];
      if (!b || (!b.top && !b.bottom && !b.left && !b.right)) {
        c++;
        continue;
      }

      // RLE: find consecutive identical borders
      let len = 1;
      while (c + len < numCols && _bordersMatch(b, borders[r][c + len])) {
        len++;
      }

      const range = sheet.getRange(minRow + r, minCol + c, 1, len);

      if (b.top)
        range.setBorder(
          true,
          null,
          null,
          null,
          null,
          null,
          b.top.color,
          b.top.style
        );
      if (b.bottom)
        range.setBorder(
          null,
          null,
          true,
          null,
          null,
          null,
          b.bottom.color,
          b.bottom.style
        );
      if (b.left)
        range.setBorder(
          null,
          true,
          null,
          null,
          null,
          null,
          b.left.color,
          b.left.style
        );
      if (b.right)
        range.setBorder(
          null,
          null,
          null,
          true,
          null,
          null,
          b.right.color,
          b.right.style
        );

      c += len;
    }
  }
}

function _bordersMatch(a, b) {
  if (a === b) return true;
  if (!a || !b) return false;
  const eq = (x, y) => x?.color === y?.color && x?.style === y?.style;
  return (
    eq(a.top, b.top) &&
    eq(a.bottom, b.bottom) &&
    eq(a.left, b.left) &&
    eq(a.right, b.right)
  );
}
