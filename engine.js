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
        
        this.resolvedCells = rootComponent.render(this, { row: 1, col: 1 }, new Style());
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
            alignment: { ...inheritedStyle.props.alignment, ...ownStyle.props.alignment },
            wrap: { ...inheritedStyle.props.wrap, ...ownStyle.props.wrap },
            border: new Border({ ...inheritedStyle.props.border.props, ...ownStyle.props.border.props }),
            rotation: { ...inheritedStyle.props.rotation, ...ownStyle.props.rotation }
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

        let minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;
        this.resolvedCells.forEach(cell => {
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

        const createGrid = (fill = null) => Array.from({ length: numRows }, () => Array(numCols).fill(fill));
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
        const borders = createGrid(false);
        const widths = {}; // Using an object to store column widths
        const heights = {}; // Using an object to store row heights
        const merges = [];
        const conditionalFormats = [];

        this.resolvedCells.forEach(cell => {
            const { props } = cell.descriptor;
            const { type, style, note, rowSpan = 1, colSpan = 1 } = props;
            const directives = type.getRenderDirectives(this.sheet.getRange(cell.row, cell.col, rowSpan, colSpan));

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
                fontWeight: style.props.font.bold ? 'bold' : 'normal',
                fontStyle: style.props.font.italic ? 'italic' : 'normal',
                fontLine: style.props.font.underline ? 'underline' : (style.props.font.strikethrough ? 'line-through' : null),
                hAlign: style.props.alignment.horizontal,
                vAlign: style.props.alignment.vertical,
                wrap: style.props.wrap.strategy,
                validation: directives.validation || null,
                border: style.props.border.props,
                numberFormat: directives.numberFormat || null
            };

            if (directives.conditionalFormatRules) {
                conditionalFormats.push(...directives.conditionalFormatRules);
            }

            for (let rOffset = 0; rOffset < rowSpan; rOffset++) {
                for (let cOffset = 0; cOffset < colSpan; cOffset++) {
                    const r = (cell.row - minRow) + rOffset;
                    const c = (cell.col - minCol) + cOffset;

                    values[r][c] = (rOffset === 0 && cOffset === 0) ? cellData.value : '';
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
            .setDataValidations(validations)
            .setNumberFormats(numberFormats)
            .setWrapStrategies(wrapStrategies);

        for (const col in widths) {
            this.sheet.setColumnWidth(parseInt(col), widths[col]);
        }
        for (const row in heights) {
            this.sheet.setRowHeight(parseInt(row), heights[row]);
        }

        for (let r = 0; r < numRows; r++) {
            for (let c = 0; c < numCols; c++) {
                const borderInfo = borders[r][c];
                if (borderInfo) {
                    const range = this.sheet.getRange(minRow + r, minCol + c);
                    const { top, bottom, left, right } = borderInfo;
                    range.setBorder(
                        top ? true : null,
                        left ? true : null,
                        bottom ? true : null,
                        right ? true : null,
                        false,
                        false,
                        top ? top.color : (right ? right.color : (bottom ? bottom.color : (left ? left.color : null))),
                        top ? top.thickness : (right ? right.thickness : (bottom ? bottom.thickness : (left ? left.thickness : null)))
                    );
                }
            }
        }

        if (conditionalFormats.length > 0) {
            this.sheet.setConditionalFormatRules(conditionalFormats);
        }

        merges.forEach(range => range.merge());
    }
}
