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
    CLIP: SpreadsheetApp.WrapStrategy.CLIP
};

const BorderThickness = {
    DOTTED: SpreadsheetApp.BorderStyle.DOTTED,
    DASHED: SpreadsheetApp.BorderStyle.DASHED,
    SOLID: SpreadsheetApp.BorderStyle.SOLID,
    SOLID_MEDIUM: SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
    SOLID_THICK: SpreadsheetApp.BorderStyle.SOLID_THICK,
    DOUBLE: SpreadsheetApp.BorderStyle.DOUBLE
};

class Component {
    constructor() {
        // The base component does not hold any data.
    }

    render() {
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
    constructor(value = '') {
        super();
        if (typeof value !== 'string') {
            throw new Error("Text component requires a string value.");
        }
        this.props = { value: value };
    }
}

class Checkbox extends Type {
    constructor(isChecked = false) {
        super();
        if (typeof isChecked !== 'boolean') {
            throw new Error("Checkbox component requires a boolean value.");
        }
        this.props = { value: isChecked };
    }

    getRenderDirectives(range) {
        return {
            validation: SpreadsheetApp.newDataValidation().requireCheckbox().build()
        };
    }
}

class Dropdown extends Type {
    constructor({ values, selected = null }) {
        super();
        if (!values || !Array.isArray(values) || values.length === 0) {
            throw new Error("Dropdown component requires a non-empty 'values' prop array.");
        }

        const isObjectArray = typeof values[0] === 'object' && values[0] !== null && 'value' in values[0];
        if (isObjectArray) {
            values.forEach(v => {
                if (!('value' in v) || !('style' in v) || !(v.style instanceof Style)) {
                    throw new Error("Dropdown 'values' array must be composed of objects with 'value' and 'style' properties, and 'style' must be an instance of Style.");
                }
            });
        }
        const plainValues = isObjectArray ? values.map(v => v.value) : values;
        const initialSelection = selected !== null ? selected : plainValues[0];

        if (selected && !plainValues.includes(selected)) {
            throw new Error(`The selected value "${selected}" is not in the list of available values.`);
        }

        this.props = {
            values: values,
            plainValues: plainValues,
            value: initialSelection,
            isObjectArray: isObjectArray
        };
    }

    getRenderDirectives(range) {
        const directives = {
            validation: SpreadsheetApp.newDataValidation().requireValueInList(this.props.plainValues).build()
        };

        if (this.props.isObjectArray) {
            const rules = [];
            this.props.values.forEach(item => {
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
    constructor({ format = '' } = {}) {
        super();
        if (typeof format !== 'string') {
            throw new Error("DatePicker format must be a string.");
        }
        this.props = { format };
    }

    getRenderDirectives(range) {
        const directives = {
            validation: SpreadsheetApp.newDataValidation().requireDate().build()
        };
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
                if (typeof side !== 'object' || side === null) {
                    throw new Error(`Border '${sideName}' must be an object.`);
                }
                if (!('color' in side) || typeof side.color !== 'string') {
                    throw new Error(`Border '${sideName}' must have a 'color' property of type string.`);
                }
                if (!('thickness' in side) || !Object.values(BorderThickness).includes(side.thickness)) {
                    throw new Error(`Border '${sideName}' must have a 'thickness' property that is a valid BorderThickness.`);
                }
            }
        };

        validateBorderSide(top, 'top');
        validateBorderSide(bottom, 'bottom');
        validateBorderSide(left, 'left');
        validateBorderSide(right, 'right');

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
        rotation = {}
    } = {}) {
        if (backgroundColor && typeof backgroundColor !== 'string') {
            throw new Error("Style 'backgroundColor' must be a string.");
        }
        if (typeof font !== 'object' || font === null) {
            throw new Error("Style 'font' must be an object.");
        }
        if (typeof alignment !== 'object' || alignment === null) {
            throw new Error("Style 'alignment' must be an object.");
        }
        if (typeof wrap !== 'object' || wrap === null) {
            throw new Error("Style 'wrap' must be an object.");
        }
        if (wrap.strategy && !Object.values(WrapStrategy).includes(wrap.strategy)) {
            throw new Error(`Invalid wrap strategy: ${wrap.strategy}`);
        }
        if (border && !(border instanceof Border)) {
            throw new Error("Style 'border' prop must be an instance of Border.");
        }
        if (typeof rotation !== 'object' || rotation === null) {
            throw new Error("Style 'rotation' must be an object.");
        }

        const defaultFont = {
            color: 'black',
            size: 10,
            family: 'Arial',
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false
        };

        const defaultAlignment = {
            horizontal: 'left',
            vertical: 'top'
        };

        const defaultWrap = {
            strategy: WrapStrategy.OVERFLOW
        };
        
        const defaultRotation = {
            angle: 0
        };

        this.props = {
            backgroundColor,
            font: { ...defaultFont, ...font },
            alignment: { ...defaultAlignment, ...alignment },
            wrap: { ...defaultWrap, ...wrap },
            border: border,
            rotation: { ...defaultRotation, ...rotation }
        };
    }
}

class HStack extends Component {
    constructor({ children, style = null }) {
        super();
        if (!children || !Array.isArray(children)) {
            throw new Error("HStack component requires a 'children' prop, which must be an array.");
        }
        if (style && !(style instanceof Style)) {
            throw new Error("HStack 'style' prop must be an instance of Style.");
        }
        this.props = { children, style };
    }

    render(position, inheritedStyle, occupancyMap) {
        let resolvedCells = [];
        let cursor = { ...position };
        const containerStyle = this.props.style ? mergeStyles(inheritedStyle, this.props.style) : inheritedStyle;

        for (const child of this.props.children) {
            if (!(child instanceof Component)) {
                throw new Error("HStack children must be instances of Component.");
            }
            let childStartPos = { ...cursor };
            while (occupancyMap.has(`${childStartPos.row}:${childStartPos.col}`)) {
                childStartPos.col++;
            }

            const childCells = child.render(childStartPos, containerStyle, occupancyMap);
            resolvedCells.push(...childCells);

            let childMaxCol = 0;
            childCells.forEach(c => {
                childMaxCol = Math.max(childMaxCol, c.col + (c.descriptor.props.colSpan || 1) - 1);
            });
            
            cursor.col = childMaxCol + 1;
        }
        return resolvedCells;
    }
}

class VStack extends Component {
    constructor({ children, style = null }) {
        super();
        if (!children || !Array.isArray(children)) {
            throw new Error("VStack component requires a 'children' prop, which must be an array.");
        }
        if (style && !(style instanceof Style)) {
            throw new Error("VStack 'style' prop must be an instance of Style.");
        }
        this.props = { children, style };
    }

    render(position, inheritedStyle, occupancyMap) {
        let resolvedCells = [];
        let cursor = { ...position };
        const containerStyle = this.props.style ? mergeStyles(inheritedStyle, this.props.style) : inheritedStyle;

        for (const child of this.props.children) {
            if (!(child instanceof Component)) {
                throw new Error("VStack children must be instances of Component.");
            }
            let childStartPos = { ...cursor };
            while (occupancyMap.has(`${childStartPos.row}:${childStartPos.col}`)) {
                childStartPos.row++;
            }

            const childCells = child.render(childStartPos, containerStyle, occupancyMap);
            resolvedCells.push(...childCells);

            let childMaxRow = 0;
            childCells.forEach(c => {
                childMaxRow = Math.max(childMaxRow, c.row + (c.descriptor.props.rowSpan || 1) - 1);
            });

            cursor.row = childMaxRow + 1;
        }
        return resolvedCells;
    }
}

class Cell extends Component {
    constructor({ type, style, note = '', colSpan = 1, rowSpan = 1 }) {
        super();
        if (type && !(type instanceof Type)) {
            throw new Error("Cell 'type' prop must be an instance of Type.");
        }
        if (style && !(style instanceof Style)) {
            throw new Error("Cell 'style' prop must be an instance of Style.");
        }
        if (typeof note !== 'string') {
            throw new Error("Cell 'note' prop must be a string.");
        }
        if (typeof colSpan !== 'number' || colSpan < 1) {
            throw new Error("Cell 'colSpan' prop must be a positive number.");
        }
        if (typeof rowSpan !== 'number' || rowSpan < 1) {
            throw new Error("Cell 'rowSpan' prop must be a positive number.");
        }
        this.props = { type, style, note, colSpan, rowSpan };
    }

    render(position, inheritedStyle, occupancyMap) {
        const { type, style, note, colSpan, rowSpan } = this.props;
        const finalType = type || new Text();
        const finalStyle = style || new Style();

        const mergedStyle = mergeStyles(inheritedStyle, finalStyle);
        
        const resolvedCell = {
            row: position.row,
            col: position.col,
            descriptor: { ...this, props: { ...this.props, style: mergedStyle } }
        };

        for (let r = 0; r < (rowSpan || 1); r++) {
            for (let c = 0; c < (colSpan || 1); c++) {
                occupancyMap.add(`${position.row + r}:${position.col + c}`);
            }
        }
        return [resolvedCell];
    }
}
