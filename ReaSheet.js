/**
 * =======================================================================
 * ReaSheets (Single-File Distribution)
 *
 * A declarative, component-based library for Google Apps Script.
 *
 * Public API is exposed globally for ease of use in GAS.
 * Internal implementation details are hidden in a closure.
 * =======================================================================
 */

/*
========================================================================
                            1. PUBLIC CONSTANTS & ENUMS
========================================================================
*/

var WrapStrategy = Object.freeze({
  WRAP: SpreadsheetApp.WrapStrategy.WRAP,
  OVERFLOW: SpreadsheetApp.WrapStrategy.OVERFLOW,
  CLIP: SpreadsheetApp.WrapStrategy.CLIP,
});

var BorderThickness = Object.freeze({
  DOTTED: SpreadsheetApp.BorderStyle.DOTTED,
  DASHED: SpreadsheetApp.BorderStyle.DASHED,
  SOLID: SpreadsheetApp.BorderStyle.SOLID,
  SOLID_MEDIUM: SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
  SOLID_THICK: SpreadsheetApp.BorderStyle.SOLID_THICK,
  DOUBLE: SpreadsheetApp.BorderStyle.DOUBLE,
});

var NumberFormats = Object.freeze({
  PERCENTAGE: "0.00%",
  CURRENCY: "$#,##0.00",
});

var FontWeight = Object.freeze({
  BOLD: "bold",
  NORMAL: "normal",
});

var FontStyle = Object.freeze({
  ITALIC: "italic",
  NORMAL: "normal",
});

var FontLine = Object.freeze({
  UNDERLINE: "underline",
  STRIKETHROUGH: "line-through",
  NONE: null,
});

// Forward declarations for public classes to be populated by the closure
var VStack,
  HStack,
  Cell,
  Text,
  Checkbox,
  Dropdown,
  DatePicker,
  NumberCell,
  Style,
  Border,
  render;

(function () {
  /*
  ========================================================================
                              2. INTERNAL CONSTANTS & UTILITIES
  ========================================================================
  */

  const InternalConstants = {
    NUMBER_FORMAT: "General",
    BORDER_SIDES: Object.freeze([
      { key: "top", args: [true, null, null, null, null, null] },
      { key: "left", args: [null, true, null, null, null, null] },
      { key: "bottom", args: [null, null, true, null, null, null] },
      { key: "right", args: [null, null, null, true, null, null] },
    ]),
  };

  const Defaults = {
    FONT: Object.freeze({
      color: "black",
      size: 10,
      family: "Arial",
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
    }),

    ALIGNMENT: Object.freeze({
      horizontal: "left",
      vertical: "top",
    }),

    WRAP: Object.freeze({
      strategy: WrapStrategy.OVERFLOW,
    }),

    ROTATION: Object.freeze({
      angle: 0,
    }),
  };

  function assertType(value, type, name, allowNull = false) {
    if (allowNull && value === null) return;
    if (type === "array") {
      if (!Array.isArray(value)) {
        throw new Error(`${name} must be an array.`);
      }
    } else if (typeof value !== type) {
      throw new Error(`${name} must be a ${type}.`);
    }
  }

  function assertInstance(value, constructor, name, allowNull = false) {
    if (allowNull && value == null) return;
    if (!(value instanceof constructor)) {
      throw new Error(`${name} must be an instance of ${constructor.name}.`);
    }
  }

  function assertPositive(value, name) {
    if (typeof value !== "number" || value < 1) {
      throw new Error(`${name} must be a positive number.`);
    }
  }

  function assertNonEmptyArray(value, name) {
    if (!Array.isArray(value) || value.length === 0) {
      throw new Error(`${name} must be a non-empty array.`);
    }
  }

  function areBordersEqual(b1, b2) {
    if (b1 === b2) return true;
    if (!b1 || !b2) return false;

    const sides = ["top", "bottom", "left", "right"];
    for (const side of sides) {
      const s1 = b1[side];
      const s2 = b2[side];
      if (s1 === s2) continue;
      if (!s1 || !s2) return false;
      if (s1.color !== s2.color || s1.thickness !== s2.thickness) return false;
    }
    return true;
  }

  /*
  ========================================================================
                              3. STYLING SYSTEM
  ========================================================================
  */

  const BORDER_THICKNESS_VALUES = Object.freeze(Object.values(BorderThickness));

  function validateBorderSide(side, sideName) {
    if (!side) return;
    assertType(side, "object", `Border '${sideName}'`);
    assertType(side.color, "string", `Border '${sideName}'.color`);
    if (!BORDER_THICKNESS_VALUES.includes(side.thickness)) {
      throw new Error(
        `Border '${sideName}'.thickness must be a valid BorderThickness.`
      );
    }
  }

  // Assign to global variable
  Border = class Border {
    constructor({ top = null, bottom = null, left = null, right = null }) {
      validateBorderSide(top, "top");
      validateBorderSide(bottom, "bottom");
      validateBorderSide(left, "left");
      validateBorderSide(right, "right");
      this.props = { top, bottom, left, right };
    }
  };

  // Assign to global variable
  Style = class Style {
    constructor({
      backgroundColor = null,
      font = {},
      alignment = {},
      wrap = {},
      border = new Border({}),
      rotation = {},
      width = null,
      height = null,
    } = {}) {
      assertType(backgroundColor, "string", "Style backgroundColor", true);
      assertType(font, "object", "Style font");
      assertType(alignment, "object", "Style alignment");
      assertType(wrap, "object", "Style wrap");
      assertType(rotation, "object", "Style rotation");
      assertType(width, "number", "Style width", true);
      assertType(height, "number", "Style height", true);
      assertInstance(border, Border, "Style border", true);

      if (
        wrap.strategy &&
        !Object.values(WrapStrategy).includes(wrap.strategy)
      ) {
        throw new Error(`Invalid wrap strategy: ${wrap.strategy}`);
      }

      this.props = {
        backgroundColor,
        font: { ...Defaults.FONT, ...font },
        alignment: { ...Defaults.ALIGNMENT, ...alignment },
        wrap: { ...Defaults.WRAP, ...wrap },
        border: border,
        rotation: { ...Defaults.ROTATION, ...rotation },
        width,
        height,
      };
    }
  };

  /*
  ========================================================================
                              4. DATA TYPES
  ========================================================================
  */

  class Type {
    constructor() {
      // Base class for type descriptors
    }

    getRenderDirectives(range) {
      return {};
    }
  }

  // Assign to global variable
  Text = class Text extends Type {
    constructor(value = "") {
      super();
      assertType(value, "string", "Text value");
      this.props = { value };
    }
  };

  // Assign to global variable
  Checkbox = class Checkbox extends Type {
    constructor(isChecked = false) {
      super();
      assertType(isChecked, "boolean", "Checkbox value");
      this.props = { value: isChecked };
    }

    getRenderDirectives(range) {
      return {
        validation: SpreadsheetApp.newDataValidation()
          .requireCheckbox()
          .build(),
      };
    }
  };

  function validateDropdownObjectArray(values) {
    values.forEach((v, i) => {
      if (!("value" in v) || !("style" in v)) {
        throw new Error(
          `Dropdown values[${i}] must have 'value' and 'style' properties.`
        );
      }
      assertInstance(v.style, Style, `Dropdown values[${i}].style`);
    });
  }

  // Assign to global variable
  Dropdown = class Dropdown extends Type {
    constructor({ values, selected = null }) {
      super();
      assertNonEmptyArray(values, "Dropdown values");

      const isObjectArray =
        typeof values[0] === "object" &&
        values[0] !== null &&
        "value" in values[0];
      if (isObjectArray) {
        validateDropdownObjectArray(values);
      }

      const plainValues = isObjectArray ? values.map((v) => v.value) : values;
      const initialSelection = selected !== null ? selected : plainValues[0];

      if (selected && !plainValues.includes(selected)) {
        throw new Error(
          `Selected value "${selected}" is not in the values list.`
        );
      }

      this.props = {
        values,
        plainValues,
        value: initialSelection,
        isObjectArray,
      };
    }

    getRenderDirectives(range) {
      const directives = {
        validation: SpreadsheetApp.newDataValidation()
          .requireValueInList(this.props.plainValues)
          .build(),
      };

      if (this.props.isObjectArray) {
        const rules = [];
        this.props.values.forEach((item) => {
          if (item.style) {
            const rule = SpreadsheetApp.newConditionalFormatRule()
              .whenTextEqualTo(item.value)
              .setBackground(item.style.props.backgroundColor)
              .setFontColor(item.style.props.font.color)
              .setRanges([range])
              .build();
            rules.push(rule);
          }
        });
        if (rules.length > 0) {
          directives.conditionalFormatRules = rules;
        }
      }
      return directives;
    }
  };

  // Assign to global variable
  DatePicker = class DatePicker extends Type {
    constructor({ format = "", value = null } = {}) {
      super();
      assertType(format, "string", "DatePicker format");
      if (value !== null) {
        assertInstance(value, Date, "DatePicker value");
        if (isNaN(value.getTime())) {
          throw new Error("DatePicker value must be a valid Date.");
        }
      }
      this.props = { format, value };
    }

    getRenderDirectives(range) {
      const directives = {
        validation: SpreadsheetApp.newDataValidation().requireDate().build(),
      };
      if (this.props.format) {
        directives.numberFormat = this.props.format;
      }
      return directives;
    }
  };

  // Assign to global variable
  NumberCell = class NumberCell extends Type {
    constructor(value, format = "0") {
      super();
      assertType(value, "number", "NumberCell value");
      assertType(format, "string", "NumberCell format");
      this.props = { value, format };
    }

    getRenderDirectives(range) {
      const directives = {};
      if (this.props.format) {
        directives.numberFormat = this.props.format;
      }
      return directives;
    }
  };

  /*
  ========================================================================
                              5. COMPONENTS
  ========================================================================
  */

  class LayoutCursor {
    constructor(renderer, startPosition) {
      this.renderer = renderer;
      this.row = startPosition.row;
      this.col = startPosition.col;
    }

    advanceToNextUnoccupied(direction) {
      while (this.renderer.occupancyMap.has(`${this.row}:${this.col}`)) {
        if (direction === "horizontal") {
          this.col++;
        } else {
          this.row++;
        }
      }
      return { row: this.row, col: this.col };
    }

    updateAfterChild(childMaxRow, childMaxCol, direction) {
      if (direction === "horizontal") {
        this.col = childMaxCol + 1;
      } else {
        this.row = childMaxRow + 1;
      }
    }
  }

  class Component {
    constructor() {
      // The base component does not hold any data.
    }

    render(renderer, position, inheritedStyle) {
      throw new Error("Component must implement a render method.");
    }
  }

  // Assign to global variable
  HStack = class HStack extends Component {
    constructor({ children, style = null }) {
      super();
      assertType(children, "array", "HStack children");
      assertInstance(style, Style, "HStack style", true);
      this.props = { children, style };
    }

    render(renderer, position, inheritedStyle) {
      let resolved = [];
      const cursor = new LayoutCursor(renderer, position);
      const containerStyle = this.props.style
        ? renderer._mergeStyles(inheritedStyle, this.props.style)
        : inheritedStyle;

      for (const child of this.props.children) {
        const childStartPos = cursor.advanceToNextUnoccupied("horizontal");
        const childCells = child.render(
          renderer,
          childStartPos,
          containerStyle
        );
        resolved.push(...childCells);

        let childMaxCol = 0;
        childCells.forEach((c) => {
          childMaxCol = Math.max(
            childMaxCol,
            c.col + (c.descriptor.props.colSpan || 1) - 1
          );
        });

        cursor.updateAfterChild(0, childMaxCol, "horizontal");
      }
      return resolved;
    }
  };

  // Assign to global variable
  VStack = class VStack extends Component {
    constructor({ children, style = null }) {
      super();
      assertType(children, "array", "VStack children");
      assertInstance(style, Style, "VStack style", true);
      this.props = { children, style };
    }

    render(renderer, position, inheritedStyle) {
      let resolved = [];
      const cursor = new LayoutCursor(renderer, position);
      const containerStyle = this.props.style
        ? renderer._mergeStyles(inheritedStyle, this.props.style)
        : inheritedStyle;

      for (const child of this.props.children) {
        const childStartPos = cursor.advanceToNextUnoccupied("vertical");
        const childCells = child.render(
          renderer,
          childStartPos,
          containerStyle
        );
        resolved.push(...childCells);

        let childMaxRow = 0;
        childCells.forEach((c) => {
          childMaxRow = Math.max(
            childMaxRow,
            c.row + (c.descriptor.props.rowSpan || 1) - 1
          );
        });

        cursor.updateAfterChild(childMaxRow, 0, "vertical");
      }
      return resolved;
    }
  };

  // Assign to global variable
  Cell = class Cell extends Component {
    constructor({
      type = new Text(""),
      style,
      note = "",
      colSpan = 1,
      rowSpan = 1,
    }) {
      super();
      assertInstance(type, Type, "Cell type", true);
      assertInstance(style, Style, "Cell style", true);
      assertType(note, "string", "Cell note");
      assertPositive(colSpan, "Cell colSpan");
      assertPositive(rowSpan, "Cell rowSpan");
      this.props = { type, style, note, colSpan, rowSpan };
    }

    render(renderer, position, inheritedStyle) {
      const { type, style, note, colSpan, rowSpan } = this.props;
      const finalType = type || new Text();
      const finalStyle = style || new Style();
      const mergedStyle = renderer._mergeStyles(inheritedStyle, finalStyle);

      const resolvedCell = {
        row: position.row,
        col: position.col,
        descriptor: { ...this, props: { ...this.props, style: mergedStyle } },
      };

      for (let r = 0; r < (rowSpan || 1); r++) {
        for (let c = 0; c < (colSpan || 1); c++) {
          renderer.occupancyMap.add(`${position.row + r}:${position.col + c}`);
        }
      }
      return [resolvedCell];
    }
  };

  /*
  ========================================================================
                              6. RENDERER ENGINE
  ========================================================================
  */

  class Renderer {
    constructor(targetSheet) {
      if (!targetSheet) {
        throw new Error("Renderer requires a target sheet object.");
      }
      this.sheet = targetSheet;
      this.occupancyMap = new Set();
      this.resolvedCells = [];
    }

    render(rootComponent) {
      if (!(rootComponent instanceof Component)) {
        throw new Error("Render function requires a root Component instance.");
      }

      this.resolvedCells = rootComponent.render(
        this,
        { row: 1, col: 1 },
        new Style()
      );
      this._commit();
    }

    _mergeStyles(inheritedStyle, ownStyle) {
      if (!inheritedStyle && !ownStyle) return new Style();
      if (!inheritedStyle) return ownStyle;
      if (!ownStyle) return inheritedStyle;

      const mergedProps = {
        ...inheritedStyle.props,
        ...ownStyle.props,
        font: { ...inheritedStyle.props.font, ...ownStyle.props.font },
        alignment: {
          ...inheritedStyle.props.alignment,
          ...ownStyle.props.alignment,
        },
        wrap: { ...inheritedStyle.props.wrap, ...ownStyle.props.wrap },
        border: new Border({
          ...inheritedStyle.props.border.props,
          ...ownStyle.props.border.props,
        }),
        rotation: {
          ...inheritedStyle.props.rotation,
          ...ownStyle.props.rotation,
        },
      };

      if (inheritedStyle.props.width !== null) {
        mergedProps.width = inheritedStyle.props.width;
      }
      if (inheritedStyle.props.height !== null) {
        mergedProps.height = inheritedStyle.props.height;
      }

      return new Style(mergedProps);
    }

    _commit() {
      if (!this.resolvedCells || this.resolvedCells.length === 0) return;

      const bounds = this._calculateBounds();
      if (!bounds) return;

      const { minRow, minCol, numRows, numCols } = bounds;
      const fullRange = this.sheet.getRange(minRow, minCol, numRows, numCols);
      fullRange.clear();

      const grids = this._buildGrids(bounds);
      this._applyStyles(fullRange, grids);
      this._applyDimensions(grids.widths, grids.heights);
      this._applyRotations(fullRange, grids.rotations);
      this._applyBorders(bounds, grids.borders);
      this._applyConditionalFormats(grids.conditionalFormats);
      this._applyMerges(grids.merges);
    }

    _calculateBounds() {
      let minRow = Infinity,
        maxRow = 0,
        minCol = Infinity,
        maxCol = 0;

      this.resolvedCells.forEach((cell) => {
        const { row, col } = cell;
        const { rowSpan = 1, colSpan = 1 } = cell.descriptor.props;
        minRow = Math.min(minRow, row);
        maxRow = Math.max(maxRow, row + rowSpan - 1);
        minCol = Math.min(minCol, col);
        maxCol = Math.max(maxCol, col + colSpan - 1);
      });

      const numRows = maxRow - minRow + 1;
      const numCols = maxCol - minCol + 1;

      if (numRows <= 0 || numCols <= 0) return null;
      return { minRow, maxRow, minCol, maxCol, numRows, numCols };
    }

    _buildGrids({ minRow, minCol, numRows, numCols }) {
      const createGrid = (fill = null) =>
        Array.from({ length: numRows }, () => Array(numCols).fill(fill));

      const grids = {
        values: createGrid(),
        backgrounds: createGrid(),
        fontColors: createGrid(),
        fontSizes: createGrid(),
        fontWeights: createGrid(),
        fontStyles: createGrid(),
        fontLines: createGrid(),
        horizontalAlignments: createGrid(),
        verticalAlignments: createGrid(),
        wrapStrategies: createGrid(WrapStrategy.OVERFLOW),
        notes: createGrid(),
        validations: createGrid(),
        numberFormats: createGrid(InternalConstants.NUMBER_FORMAT),
        rotations: createGrid(0),
        borders: createGrid(false),
        widths: {},
        heights: {},
        merges: [],
        conditionalFormats: [],
      };

      this.resolvedCells.forEach((cell) => {
        this._populateGridsForCell(cell, grids, minRow, minCol);
      });

      return grids;
    }

    _populateGridsForCell(cell, grids, minRow, minCol) {
      const { props } = cell.descriptor;
      const { type, style, note, rowSpan = 1, colSpan = 1 } = props;
      const directives = type.getRenderDirectives(
        this.sheet.getRange(cell.row, cell.col, rowSpan, colSpan)
      );

      if (style.props.width !== null)
        grids.widths[cell.col] = style.props.width;
      if (style.props.height !== null)
        grids.heights[cell.row] = style.props.height;

      const cellData = this._extractCellData(type, style, note, directives);

      if (directives.conditionalFormatRules) {
        grids.conditionalFormats.push(...directives.conditionalFormatRules);
      }

      for (let rOffset = 0; rOffset < rowSpan; rOffset++) {
        for (let cOffset = 0; cOffset < colSpan; cOffset++) {
          const r = cell.row - minRow + rOffset;
          const c = cell.col - minCol + cOffset;
          this._fillGridCell(
            grids,
            r,
            c,
            cellData,
            rOffset,
            cOffset,
            rowSpan,
            colSpan
          );
        }
      }

      if (rowSpan > 1 || colSpan > 1) {
        grids.merges.push(
          this.sheet.getRange(cell.row, cell.col, rowSpan, colSpan)
        );
      }
    }

    _extractCellData(type, style, note, directives) {
      const { font, alignment, wrap, border, rotation } = style.props;
      return {
        value: type.props.value,
        note,
        background: style.props.backgroundColor,
        fontColor: font.color,
        fontSize: font.size,
        fontWeight: font.bold ? FontWeight.BOLD : FontWeight.NORMAL,
        fontStyle: font.italic ? FontStyle.ITALIC : FontStyle.NORMAL,
        fontLine: font.underline
          ? FontLine.UNDERLINE
          : font.strikethrough
          ? FontLine.STRIKETHROUGH
          : FontLine.NONE,
        hAlign: alignment.horizontal,
        vAlign: alignment.vertical,
        wrap: wrap.strategy,
        validation: directives.validation || null,
        border: border.props,
        numberFormat: directives.numberFormat || null,
        rotation: rotation.angle,
      };
    }

    _fillGridCell(grids, r, c, cellData, rOffset, cOffset, rowSpan, colSpan) {
      grids.values[r][c] = rOffset === 0 && cOffset === 0 ? cellData.value : "";
      grids.notes[r][c] = cellData.note;
      grids.backgrounds[r][c] = cellData.background;
      grids.fontColors[r][c] = cellData.fontColor;
      grids.fontSizes[r][c] = cellData.fontSize;
      grids.fontWeights[r][c] = cellData.fontWeight;
      grids.fontStyles[r][c] = cellData.fontStyle;
      grids.fontLines[r][c] = cellData.fontLine;
      grids.horizontalAlignments[r][c] = cellData.hAlign;
      grids.verticalAlignments[r][c] = cellData.vAlign;
      grids.wrapStrategies[r][c] = cellData.wrap;
      grids.validations[r][c] = cellData.validation;
      grids.rotations[r][c] = cellData.rotation;
      if (cellData.numberFormat)
        grids.numberFormats[r][c] = cellData.numberFormat;

      const { top, bottom, left, right } = cellData.border;
      grids.borders[r][c] = {
        top: rOffset === 0 && top,
        bottom: rOffset === rowSpan - 1 && bottom,
        left: cOffset === 0 && left,
        right: cOffset === colSpan - 1 && right,
      };
    }

    _applyStyles(range, grids) {
      range
        .setNumberFormats(grids.numberFormats)
        .setDataValidations(grids.validations)
        .setValues(grids.values)
        .setNotes(grids.notes)
        .setBackgrounds(grids.backgrounds)
        .setFontColors(grids.fontColors)
        .setFontSizes(grids.fontSizes)
        .setFontWeights(grids.fontWeights)
        .setFontStyles(grids.fontStyles)
        .setFontLines(grids.fontLines)
        .setHorizontalAlignments(grids.horizontalAlignments)
        .setVerticalAlignments(grids.verticalAlignments)
        .setWrapStrategies(grids.wrapStrategies);
    }

    _applyDimensions(widths, heights) {
      for (const col in widths) {
        this.sheet.setColumnWidth(parseInt(col), widths[col]);
      }
      for (const row in heights) {
        this.sheet.setRowHeight(parseInt(row), heights[row]);
      }
    }

    _applyRotations(range, rotations) {
      const hasRotation = rotations.some((row) =>
        row.some((angle) => angle !== 0)
      );
      if (hasRotation) {
        range.setTextRotations(rotations);
      }
    }

    _applyBorders({ minRow, minCol, numRows, numCols }, borders) {
      for (let r = 0; r < numRows; r++) {
        let c = 0;
        while (c < numCols) {
          const borderInfo = borders[r][c];
          if (!borderInfo) {
            c++;
            continue;
          }

          let len = 1;
          while (
            c + len < numCols &&
            areBordersEqual(borderInfo, borders[r][c + len])
          ) {
            len++;
          }

          const range = this.sheet.getRange(minRow + r, minCol + c, 1, len);

          if (borderInfo.top) {
            range.setBorder(
              true,
              null,
              null,
              null,
              null,
              null,
              borderInfo.top.color,
              borderInfo.top.thickness
            );
          }
          if (borderInfo.bottom) {
            range.setBorder(
              null,
              null,
              true,
              null,
              null,
              null,
              borderInfo.bottom.color,
              borderInfo.bottom.thickness
            );
          }
          if (borderInfo.left) {
            range.setBorder(
              null,
              true,
              null,
              null,
              null,
              null,
              borderInfo.left.color,
              borderInfo.left.thickness
            );
          }
          if (borderInfo.right) {
            range.setBorder(
              null,
              null,
              null,
              true,
              null,
              null,
              borderInfo.right.color,
              borderInfo.right.thickness
            );
          }
          if (len > 1) {
            const left = borderInfo.left;
            const right = borderInfo.right;
            if (
              left &&
              right &&
              left.color === right.color &&
              left.thickness === right.thickness
            ) {
              range.setBorder(
                null,
                null,
                null,
                null,
                true,
                null,
                left.color,
                left.thickness
              );
            }
          }
          c += len;
        }
      }
    }

    _applyConditionalFormats(conditionalFormats) {
      if (conditionalFormats.length > 0) {
        this.sheet.setConditionalFormatRules(conditionalFormats);
      }
    }

    _applyMerges(merges) {
      merges.forEach((range) => range.merge());
    }
  }

  // Assign to global variable
  render = function (rootComponent, targetSheet) {
    new Renderer(targetSheet).render(rootComponent);
  };
})();