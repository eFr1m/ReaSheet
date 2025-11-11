/**
 * =======================================================================
 * ReaSheets Core Component Library
 *
 * This file contains the fundamental building blocks for the declarative
 * UI framework for Google Sheets.
 * =======================================================================
 */

const WrapStrategy = {
  WRAP: SpreadsheetApp.WrapStrategy.WRAP,
  OVERFLOW: SpreadsheetApp.WrapStrategy.OVERFLOW,
  CLIP: SpreadsheetApp.WrapStrategy.CLIP,
};

const BorderThickness = {
  DOTTED: SpreadsheetApp.BorderStyle.DOTTED,
  DASHED: SpreadsheetApp.BorderStyle.DASHED,
  SOLID: SpreadsheetApp.BorderStyle.SOLID,
  SOLID_MEDIUM: SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
  SOLID_THICK: SpreadsheetApp.BorderStyle.SOLID_THICK,
  DOUBLE: SpreadsheetApp.BorderStyle.DOUBLE,
};

const NumberFormats = {
  PERCENTAGE: "0.00%",
  CURRENCY: "$#,##0.00",
};

class Component {
  constructor() {
    // The base component does not hold any data.
  }

  render(renderer, position, inheritedStyle) {
    throw new Error("Component must implement a render method.");
  }
}

class Type {
  constructor() {
    // Base class for type descriptors
  }

  getRenderDirectives(range) {
    return {};
  }
}

class Text extends Type {
  constructor(value = "") {
    super();
    if (typeof value !== "string") {
      throw new Error("Text component requires a string value.");
    }
    this.props = { value: value };
  }
}

class Checkbox extends Type {
  constructor(isChecked = false) {
    super();
    if (typeof isChecked !== "boolean") {
      throw new Error("Checkbox component requires a boolean value.");
    }
    this.props = { value: isChecked };
  }

  getRenderDirectives(range) {
    return {
      validation: SpreadsheetApp.newDataValidation().requireCheckbox().build(),
    };
  }
}

class Dropdown extends Type {
  constructor({ values, selected = null }) {
    super();
    if (!values || !Array.isArray(values) || values.length === 0) {
      throw new Error(
        "Dropdown component requires a non-empty 'values' prop array."
      );
    }

    const isObjectArray =
      typeof values[0] === "object" &&
      values[0] !== null &&
      "value" in values[0];
    if (isObjectArray) {
      values.forEach((v) => {
        if (!("value" in v) || !("style" in v) || !(v.style instanceof Style)) {
          throw new Error(
            "Dropdown 'values' array must be composed of objects with 'value' and 'style' properties, and 'style' must be an instance of Style."
          );
        }
      });
    }
    const plainValues = isObjectArray ? values.map((v) => v.value) : values;
    const initialSelection = selected !== null ? selected : plainValues[0];

    if (selected && !plainValues.includes(selected)) {
      throw new Error(
        `The selected value "${selected}" is not in the list of available values.`
      );
    }

    this.props = {
      values: values,
      plainValues: plainValues,
      value: initialSelection,
      isObjectArray: isObjectArray,
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
}

class DatePicker extends Type {
  constructor({ format = "", value = null } = {}) {
    super();
    if (typeof format !== "string") {
      throw new Error("DatePicker format must be a string.");
    }
    if (value && (!(value instanceof Date) || isNaN(value.getTime()))) {
      throw new Error("DatePicker 'value' prop must be a valid Date object.");
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
}

class Number extends Type {
  constructor(value, format = "0") {
    super();
    if (typeof value !== "number") {
      throw new Error("Number component requires a number value.");
    }
    if (format && typeof format !== "string") {
      throw new Error("format must be a string.");
    }
    this.props = { value, format };
  }

  getRenderDirectives(range) {
    const directives = {};
    if (this.props.format) {
      directives.numberFormat = this.props.format;
    }
    return directives;
  }
}

class Border {
  constructor({ top = null, bottom = null, left = null, right = null }) {
    const validateBorderSide = (side, sideName) => {
      if (side) {
        if (typeof side !== "object" || side === null) {
          throw new Error(`Border '${sideName}' must be an object.`);
        }
        if (!("color" in side) || typeof side.color !== "string") {
          throw new Error(
            `Border '${sideName}' must have a 'color' property of type string.`
          );
        }
        if (
          !("thickness" in side) ||
          !Object.values(BorderThickness).includes(side.thickness)
        ) {
          throw new Error(
            `Border '${sideName}' must have a 'thickness' property that is a valid BorderThickness.`
          );
        }
      }
    };

    validateBorderSide(top, "top");
    validateBorderSide(bottom, "bottom");
    validateBorderSide(left, "left");
    validateBorderSide(right, "right");

    this.props = { top, bottom, left, right };
  }
}

class Style {
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
    if (backgroundColor && typeof backgroundColor !== "string") {
      throw new Error("Style 'backgroundColor' must be a string.");
    }
    if (typeof font !== "object" || font === null) {
      throw new Error("Style 'font' must be an object.");
    }
    if (typeof alignment !== "object" || alignment === null) {
      throw new Error("Style 'alignment' must be an object.");
    }
    if (typeof wrap !== "object" || wrap === null) {
      throw new Error("Style 'wrap' must be an object.");
    }
    if (wrap.strategy && !Object.values(WrapStrategy).includes(wrap.strategy)) {
      throw new Error(`Invalid wrap strategy: ${wrap.strategy}`);
    }
    if (border && !(border instanceof Border)) {
      throw new Error("Style 'border' prop must be an instance of Border.");
    }
    if (typeof rotation !== "object" || rotation === null) {
      throw new Error("Style 'rotation' must be an object.");
    }
    if (width !== null && typeof width !== "number") {
      throw new Error("Style 'width' must be a number.");
    }
    if (height !== null && typeof height !== "number") {
      throw new Error("Style 'height' must be a number.");
    }

    const defaultFont = {
      color: "black",
      size: 10,
      family: "Arial",
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
    };

    const defaultAlignment = {
      horizontal: "left",
      vertical: "top",
    };

    const defaultWrap = {
      strategy: WrapStrategy.OVERFLOW,
    };

    const defaultRotation = {
      angle: 0,
    };

    this.props = {
      backgroundColor,
      font: { ...defaultFont, ...font },
      alignment: { ...defaultAlignment, ...alignment },
      wrap: { ...defaultWrap, ...wrap },
      border: border,
      rotation: { ...defaultRotation, ...rotation },
      width,
      height,
    };
  }
}

class HStack extends Component {
  constructor({ children, style = null }) {
    super();
    if (!children || !Array.isArray(children)) {
      throw new Error(
        "HStack component requires a 'children' prop, which must be an array."
      );
    }
    if (style && !(style instanceof Style)) {
      throw new Error("HStack 'style' prop must be an instance of Style.");
    }
    this.props = { children, style };
  }

  render(renderer, position, inheritedStyle) {
    let resolved = [];
    let cursor = { ...position };
    const containerStyle = this.props.style
      ? renderer._mergeStyles(inheritedStyle, this.props.style)
      : inheritedStyle;

    for (const child of this.props.children) {
      let childStartPos = { ...cursor };
      while (
        renderer.occupancyMap.has(`${childStartPos.row}:${childStartPos.col}`)
      ) {
        childStartPos.col++;
      }

      const childCells = child.render(renderer, childStartPos, containerStyle);
      resolved.push(...childCells);

      let childMaxCol = 0;
      childCells.forEach((c) => {
        childMaxCol = Math.max(
          childMaxCol,
          c.col + (c.descriptor.props.colSpan || 1) - 1
        );
      });

      cursor.col = childMaxCol + 1;
    }
    return resolved;
  }
}

class VStack extends Component {
  constructor({ children, style = null }) {
    super();
    if (!children || !Array.isArray(children)) {
      throw new Error(
        "VStack component requires a 'children' prop, which must be an array."
      );
    }
    if (style && !(style instanceof Style)) {
      throw new Error("VStack 'style' prop must be an instance of Style.");
    }
    this.props = { children, style };
  }

  render(renderer, position, inheritedStyle) {
    let resolved = [];
    let cursor = { ...position };
    const containerStyle = this.props.style
      ? renderer._mergeStyles(inheritedStyle, this.props.style)
      : inheritedStyle;

    for (const child of this.props.children) {
      let childStartPos = { ...cursor };
      while (
        renderer.occupancyMap.has(`${childStartPos.row}:${childStartPos.col}`)
      ) {
        childStartPos.row++;
      }

      const childCells = child.render(renderer, childStartPos, containerStyle);
      resolved.push(...childCells);

      let childMaxRow = 0;
      childCells.forEach((c) => {
        childMaxRow = Math.max(
          childMaxRow,
          c.row + (c.descriptor.props.rowSpan || 1) - 1
        );
      });

      cursor.row = childMaxRow + 1;
    }
    return resolved;
  }
}

class Cell extends Component {
  constructor({
    type = new Text(""),
    style,
    note = "",
    colSpan = 1,
    rowSpan = 1,
  }) {
    super();
    if (type && !(type instanceof Type)) {
      throw new Error("Cell 'type' prop must be an instance of Type.");
    }
    if (style && !(style instanceof Style)) {
      throw new Error("Cell 'style' prop must be an instance of Style.");
    }
    if (typeof note !== "string") {
      throw new Error("Cell 'note' prop must be a string.");
    }
    if (typeof colSpan !== "number" || colSpan < 1) {
      throw new Error("Cell 'colSpan' prop must be a positive number.");
    }
    if (typeof rowSpan !== "number" || rowSpan < 1) {
      throw new Error("Cell 'rowSpan' prop must be a positive number.");
    }
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
}

/*
========================================================================
                  Renderer Class and Render Function
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

    // Parent's width and height win
    if (inheritedStyle.props.width !== null) {
      mergedProps.width = inheritedStyle.props.width;
    }
    if (inheritedStyle.props.height !== null) {
      mergedProps.height = inheritedStyle.props.height;
    }

    return new Style(mergedProps);
  }

  _commit() {
    if (!this.resolvedCells || this.resolvedCells.length === 0) {
      return;
    }

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

    if (numRows <= 0 || numCols <= 0) return;

    const fullRange = this.sheet.getRange(minRow, minCol, numRows, numCols);
    fullRange.clear();

    const createGrid = (fill = null) =>
      Array.from({ length: numRows }, () => Array(numCols).fill(fill));
    const values = createGrid();
    const backgrounds = createGrid();
    const fontColors = createGrid();
    const fontSizes = createGrid();
    const fontWeights = createGrid();
    const fontStyles = createGrid();
    const fontLines = createGrid();
    const horizontalAlignments = createGrid();
    const verticalAlignments = createGrid();
    const wrapStrategies = createGrid(WrapStrategy.OVERFLOW);
    const notes = createGrid();
    const validations = createGrid();
    const numberFormats = createGrid("General");
    const rotations = createGrid(0);
    const borders = createGrid(false);
    const widths = {}; // Using an object to store column widths
    const heights = {}; // Using an object to store row heights
    const merges = [];
    const conditionalFormats = [];

    this.resolvedCells.forEach((cell) => {
      const { props } = cell.descriptor;
      const { type, style, note, rowSpan = 1, colSpan = 1 } = props;
      const directives = type.getRenderDirectives(
        this.sheet.getRange(cell.row, cell.col, rowSpan, colSpan)
      );

      if (style.props.width !== null) {
        widths[cell.col] = style.props.width;
      }
      if (style.props.height !== null) {
        heights[cell.row] = style.props.height;
      }

      const cellData = {
        value: type.props.value,
        note: note,
        background: style.props.backgroundColor,
        fontColor: style.props.font.color,
        fontSize: style.props.font.size,
        fontWeight: style.props.font.bold ? "bold" : "normal",
        fontStyle: style.props.font.italic ? "italic" : "normal",
        fontLine: style.props.font.underline
          ? "underline"
          : style.props.font.strikethrough
          ? "line-through"
          : null,
        hAlign: style.props.alignment.horizontal,
        vAlign: style.props.alignment.vertical,
        wrap: style.props.wrap.strategy,
        validation: directives.validation || null,
        border: style.props.border.props,
        numberFormat: directives.numberFormat || null,
        rotation: style.props.rotation.angle,
      };

      if (directives.conditionalFormatRules) {
        conditionalFormats.push(...directives.conditionalFormatRules);
      }

      for (let rOffset = 0; rOffset < rowSpan; rOffset++) {
        for (let cOffset = 0; cOffset < colSpan; cOffset++) {
          const r = cell.row - minRow + rOffset;
          const c = cell.col - minCol + cOffset;

          values[r][c] = rOffset === 0 && cOffset === 0 ? cellData.value : "";
          notes[r][c] = cellData.note;
          backgrounds[r][c] = cellData.background;
          fontColors[r][c] = cellData.fontColor;
          fontSizes[r][c] = cellData.fontSize;
          fontWeights[r][c] = cellData.fontWeight;
          fontStyles[r][c] = cellData.fontStyle;
          fontLines[r][c] = cellData.fontLine;
          horizontalAlignments[r][c] = cellData.hAlign;
          verticalAlignments[r][c] = cellData.vAlign;
          wrapStrategies[r][c] = cellData.wrap;
          validations[r][c] = cellData.validation;
          rotations[r][c] = cellData.rotation;
          if (cellData.numberFormat) {
            numberFormats[r][c] = cellData.numberFormat;
          }

          const { top, bottom, left, right } = cellData.border;
          borders[r][c] = {
            top: rOffset === 0 && top,
            bottom: rOffset === rowSpan - 1 && bottom,
            left: cOffset === 0 && left,
            right: cOffset === colSpan - 1 && right,
          };
        }
      }

      if (rowSpan > 1 || colSpan > 1) {
        merges.push(this.sheet.getRange(cell.row, cell.col, rowSpan, colSpan));
      }
    });

    fullRange
      .setNumberFormats(numberFormats)
      .setDataValidations(validations)
      .setValues(values)
      .setNotes(notes)
      .setBackgrounds(backgrounds)
      .setFontColors(fontColors)
      .setFontSizes(fontSizes)
      .setFontWeights(fontWeights)
      .setFontStyles(fontStyles)
      .setFontLines(fontLines)
      .setHorizontalAlignments(horizontalAlignments)
      .setVerticalAlignments(verticalAlignments)
      .setWrapStrategies(wrapStrategies);

    for (const col in widths) {
      this.sheet.setColumnWidth(parseInt(col), widths[col]);
    }
    for (const row in heights) {
      this.sheet.setRowHeight(parseInt(row), heights[row]);
    }

    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const rotation = rotations[r][c];
        if (rotation !== 0) {
          this.sheet.getRange(minRow + r, minCol + c).setTextRotation(rotation);
        }

        const borderInfo = borders[r][c];
        if (borderInfo) {
          const range = this.sheet.getRange(minRow + r, minCol + c);
          const { top, bottom, left, right } = borderInfo;
          if (top) {
            range.setBorder(
              true,
              null,
              null,
              null,
              null,
              null,
              top.color,
              top.thickness
            );
          }
          if (left) {
            range.setBorder(
              null,
              true,
              null,
              null,
              null,
              null,
              left.color,
              left.thickness
            );
          }
          if (bottom) {
            range.setBorder(
              null,
              null,
              true,
              null,
              null,
              null,
              bottom.color,
              bottom.thickness
            );
          }
          if (right) {
            range.setBorder(
              null,
              null,
              null,
              true,
              null,
              null,
              right.color,
              right.thickness
            );
          }
        }
      }
    }

    if (conditionalFormats.length > 0) {
      this.sheet.setConditionalFormatRules(conditionalFormats);
    }

    merges.forEach((range) => range.merge());
  }
}

function render(rootComponent, targetSheet) {
  new Renderer(targetSheet).render(rootComponent);
}
