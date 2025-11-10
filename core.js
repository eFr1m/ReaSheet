/**
 * =======================================================================
 * ReaSheets Core Component Library
 *
 * This file contains the fundamental building blocks for the declarative
 * UI framework for Google Sheets.
 * =======================================================================
 */

/**
 * @description An enum-like object to define the different categories of components.
 * This is used for runtime type safety to ensure components are used in the correct props.
 * @readonly
 * @enum {string}
 */
const Category = Object.freeze({
  LAYOUT: "Layout",
  TYPE: "Type",
  STYLE: "Style",
});

const WrapStrategy = Object.freeze({
  WRAP: SpreadsheetApp.WrapStrategy.WRAP,
  OVERFLOW: SpreadsheetApp.WrapStrategy.OVERFLOW,
  CLIP: SpreadsheetApp.WrapStrategy.CLIP,
});

/**
 * Defines a standard text type.
 * @param {any} [value=''] - The value to be displayed.
 * @returns {object} A Text type descriptor.
 */
function Text(value = "") {
  return {
    category: Category.TYPE,
    component: Text,
    props: { value: String(value) },
  };
}

/**
 * Defines a checkbox type.
 * @param {boolean} [isChecked=false] - The initial state of the checkbox.
 * @returns {object} A Checkbox type descriptor.
 */
function Checkbox(isChecked = false) {
  return {
    category: Category.TYPE,
    component: Checkbox,
    props: { value: isChecked },
  };
}

/**
 * Defines a dropdown type from a list of values.
 * @param {object} props - The properties for the dropdown.
 * @param {Array<string|object>} props.values - The list of values for the dropdown. Can be strings or objects with 'value' and 'style'.
 * @param {any} [props.selected=null] - The pre-selected value. If null, defaults to the first value.
 * @returns {object} A Dropdown type descriptor.
 */
function Dropdown({ values, selected = null }) {
    if (!values || !Array.isArray(values) || values.length === 0) {
        throw new Error("Dropdown component requires a non-empty 'values' prop array.");
    }

    const isObjectArray = typeof values[0] === 'object' && values[0] !== null && 'value' in values[0];
    const plainValues = isObjectArray ? values.map(v => v.value) : values;
    const initialSelection = selected !== null ? selected : plainValues[0];

    return {
        category: Category.TYPE,
        component: Dropdown,
        props: {
            values: values, // Keep original structure for the engine
            plainValues: plainValues,
            value: initialSelection,
            isObjectArray: isObjectArray
        }
    };
}

/**
 * Defines a type that requires a cell to contain a valid date.
 * @param {object} [props={}] - The properties for the date picker.
 * @param {string} [props.format=''] - The date format string (e.g., 'yyyy-mm-dd').
 * @returns {object} A DatePicker type descriptor.
 */
function DatePicker({ format = '' } = {}) {
  return {
    category: Category.TYPE,
    component: DatePicker,
    props: { format },
  };
}

/**
 * Defines styling for a Cell.
 * The properties are designed to mirror the Google Apps Script `Range` styling options.
 * This component uses the "Safe Merge" pattern to robustly handle defaults for nested objects.
 * @param {object} [props={}] - The properties for the style.
 * @param {string} [props.backgroundColor=null] - The background color (e.g., '#ffffff').
 * @param {object} [props.font={}] - Font properties.
 * @param {object} [props.alignment={}] - Alignment properties.
 * @param {object} [props.wrap={}] - Text wrap properties.
 * @param {object} [props.border={}] - Border properties.
 * @param {object} [props.rotation={}] - Text rotation properties.
 * @returns {object} A Style component descriptor.
 */
function Style({
  backgroundColor = null,
  font = {},
  alignment = {},
  wrap = {},
  border = {},
  rotation = {},
} = {}) {
  if (wrap.strategy && !Object.values(WrapStrategy).includes(wrap.strategy)) {
    throw new Error(
      `Invalid wrap strategy provided in Style component. Received: ${wrap.strategy}`
    );
  }

  const defaultFont = {
    color: "#000000",
    size: 12,
    family: "Arial",
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
  };

  const defaultAlignment = {
    horizontal: true,
    vertical: true,
  };

  const defaultWrap = {
    strategy: WrapStrategy.OVERFLOW,
  };

  const defaultRotation = {
    angle: 0,
  };

  const finalProps = {
    backgroundColor,
    font: { ...defaultFont, ...font },
    alignment: { ...defaultAlignment, ...alignment },
    wrap: { ...defaultWrap, ...wrap },
    border,
    rotation: { ...defaultRotation, ...rotation },
  };

  return {
    category: Category.STYLE,
    component: Style,
    props: finalProps,
  };
}

/**
 * A layout component that arranges its children horizontally.
 * @param {object} props - The properties for the stack.
 * @param {Array<object>} props.children - An array of layout component descriptors (e.g., Cell, VStack).
 * @param {object} [props.style=null] - A Style component to apply to the container.
 * @returns {object} An HStack layout descriptor.
 */
function HStack({ children, style = null }) {
  if (!children || !Array.isArray(children)) {
    throw new Error(
      "HStack component requires a 'children' prop, which must be an array."
    );
  }
  for (const child of children) {
    if (!child || child.category !== Category.LAYOUT) {
      throw new Error(
        `Invalid child passed to HStack. All children must be Layout Components (e.g., Cell, VStack). Received: ${child.category}`
      );
    }
  }
  if (style && style.category !== Category.STYLE) {
    throw new Error(
      `Invalid component passed to HStack 'style' prop. Expected a Style Component.`
    );
  }

  return {
    category: Category.LAYOUT,
    component: HStack,
    props: { children, style },
  };
}

/**
 * A layout component that arranges its children vertically.
 * @param {object} props - The properties for the stack.
 * @param {Array<object>} props.children - An array of layout component descriptors (e.g., Cell, HStack).
 * @param {object} [props.style=null] - A Style component to apply to the container.
 * @returns {object} A VStack layout descriptor.
 */
function VStack({ children, style = null }) {
  if (!children || !Array.isArray(children)) {
    throw new Error(
      "VStack component requires a 'children' prop, which must be an array."
    );
  }
  for (const child of children) {
    if (!child || child.category !== Category.LAYOUT) {
      throw new Error(
        `Invalid child passed to VStack. All children must be Layout Components (e.g., Cell, HStack). Received: ${child.category}`
      );
    }
  }
  if (style && style.category !== Category.STYLE) {
    throw new Error(
      `Invalid component passed to VStack 'style' prop. Expected a Style Component.`
    );
  }

  return {
    category: Category.LAYOUT,
    component: VStack,
    props: { children, style },
  };
}

// ========================================================================
// PRIMITIVE LAYOUT COMPONENT
// This defines the physical cell container on the grid.
// ========================================================================

/**
 * Defines a single cell in the sheet. This is the primary layout component.
 * It acts as a container for a 'type' component and a 'style' component.
 *
 * @param {object} [props={}] - The properties for the cell.
 * @param {object} [props.type] - The type component defining the cell's content (e.g., Text(), Checkbox()). Defaults to an empty Text type.
 * @param {object} [props.style] - The style component for the cell. Defaults to an empty Style component.
 * @param {string} [props.note=''] - A note to attach to the cell.
 * @param {number} [props.colSpan=1] - The number of columns the cell should span.
 * @param {number} [props.rowSpan=1] - The number of rows the cell should span.
 * @returns {object} A Cell layout descriptor.
 */
function Cell({ type, style, note = "", colSpan = 1, rowSpan = 1 } = {}) {
  const finalType = type || Text();
  if (finalType.category !== Category.TYPE) {
    throw new Error(
      `Invalid descriptor passed to 'type' prop. Expected a ${Category.TYPE} descriptor (e.g., Text(), Checkbox()), but received a ${finalType.category} descriptor.`
    );
  }

  const finalStyle = style || Style();
  if (finalStyle.category !== Category.STYLE) {
    throw new Error(
      `Invalid descriptor passed to 'style' prop. Expected a ${Category.STYLE} descriptor, but received a ${finalStyle.category} descriptor.`
    );
  }

  return {
    category: Category.LAYOUT,
    component: Cell,
    props: { type: finalType, style: finalStyle, note, colSpan, rowSpan },
  };
}
