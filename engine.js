/**
 * Merges an inherited style with a component's own style.
 * The ownStyle takes precedence in case of conflicts.
 * @param {object} inheritedStyle - The style descriptor from the parent container.
 * @param {object} ownStyle - The style descriptor from the component itself.
 * @returns {object} A new, merged Style component descriptor.
 */
function mergeStyles(inheritedStyle, ownStyle) {
    if (!inheritedStyle && !ownStyle) return Style();
    if (!inheritedStyle) return ownStyle;
    if (!ownStyle) return inheritedStyle;

    const mergedProps = {
        ...inheritedStyle.props,
        ...ownStyle.props,
        font: { ...inheritedStyle.props.font, ...ownStyle.props.font },
        alignment: { ...inheritedStyle.props.alignment, ...ownStyle.props.alignment },
        wrap: { ...inheritedStyle.props.wrap, ...ownStyle.props.wrap },
        border: { ...inheritedStyle.props.border, ...ownStyle.props.border },
        rotation: { ...inheritedStyle.props.rotation, ...ownStyle.props.rotation }
    };
    return Style(mergedProps);
}

function layoutCell(descriptor, position, inheritedStyle, occupancyMap) {
    const finalStyle = mergeStyles(inheritedStyle, descriptor.props.style);
    
    const finalDescriptor = { ...descriptor };
    finalDescriptor.props.style = finalStyle;

    const resolvedCell = {
        row: position.row,
        col: position.col,
        descriptor: finalDescriptor
    };

    const { rowSpan = 1, colSpan = 1 } = descriptor.props;
    for (let r = 0; r < rowSpan; r++) {
        for (let c = 0; c < colSpan; c++) {
            occupancyMap.add(`${position.row + r}:${position.col + c}`);
        }
    }
    return [resolvedCell];
}

function layoutHStack(descriptor, position, inheritedStyle, occupancyMap) {
    let resolvedCells = [];
    let cursor = { ...position };
    const containerStyle = mergeStyles(inheritedStyle, descriptor.props.style);

    for (const child of descriptor.props.children) {
        let childStartPos = { ...cursor };
        while (occupancyMap.has(`${childStartPos.row}:${childStartPos.col}`)) {
            childStartPos.col++;
        }

        const childCells = layoutEngine(child, childStartPos, containerStyle, occupancyMap);
        resolvedCells.push(...childCells);

        let childMaxCol = 0;
        childCells.forEach(c => {
            childMaxCol = Math.max(childMaxCol, c.col + (c.descriptor.props.colSpan || 1) - 1);
        });
        
        cursor.col = childMaxCol + 1;
    }
    return resolvedCells;
}

function layoutVStack(descriptor, position, inheritedStyle, occupancyMap) {
    let resolvedCells = [];
    let cursor = { ...position };
    const containerStyle = mergeStyles(inheritedStyle, descriptor.props.style);

    for (const child of descriptor.props.children) {
        let childStartPos = { ...cursor };
        while (occupancyMap.has(`${childStartPos.row}:${childStartPos.col}`)) {
            childStartPos.row++;
        }

        const childCells = layoutEngine(child, childStartPos, containerStyle, occupancyMap);
        resolvedCells.push(...childCells);

        let childMaxRow = 0;
        childCells.forEach(c => {
            childMaxRow = Math.max(childMaxRow, c.row + (c.descriptor.props.rowSpan || 1) - 1);
        });

        cursor.row = childMaxRow + 1;
    }
    return resolvedCells;
}

/**
 * The Layout Engine.
 * Recursively walks a component descriptor tree and calculates the absolute
 * position of every cell, handling spans and style inheritance.
 * @param {object} descriptor - The root layout component descriptor.
 * @param {object} startPosition - The starting {row, col}.
 * @param {object} inheritedStyle - The style from parent containers.
 * @param {Set<string>} occupancyMap - A map of occupied "row:col" coordinates.
 * @returns {Array<object>} A flat array of resolved cells.
 */
function layoutEngine(descriptor, startPosition = { row: 1, col: 1 }, inheritedStyle = Style(), occupancyMap = new Set()) {
    if (!descriptor || descriptor.category !== Category.LAYOUT) {
        throw new Error("Layout engine can only process Layout components.");
    }

    const layoutFunctions = {
        [Cell.name]: layoutCell,
        [HStack.name]: layoutHStack,
        [VStack.name]: layoutVStack,
    };

    const layoutFunction = layoutFunctions[descriptor.component.name];

    if (layoutFunction) {
        return layoutFunction(descriptor, startPosition, inheritedStyle, occupancyMap);
    }

    return [];
}

/**
 * The Commit Engine.
 * Takes a flat array of resolved cells and applies them to the Google Sheet
 * using efficient, batched API calls.
 * @param {Array<object>} resolvedCells - The output from the layoutEngine.
 * @param {Sheet} targetSheet - The Google Apps Script sheet object to render to.
 */
function commitEngine(resolvedCells, targetSheet) {
    if (!resolvedCells || resolvedCells.length === 0) {
        return; // Nothing to render
    }

    // 1. Find Bounding Box
    let minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;
    resolvedCells.forEach(cell => {
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
    
    const fullRange = targetSheet.getRange(minRow, minCol, numRows, numCols);

    // Clear previous content and formatting
    fullRange.clear();

    // 2. Create Grids
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
    const wrapStrategies = createGrid();
    const notes = createGrid();
    const validations = createGrid();
    const numberFormats = createGrid();
    const borders = createGrid(false); // For border settings
    const merges = [];

    // 3. Populate Grids
    resolvedCells.forEach(cell => {
        const { props } = cell.descriptor;
        const { type, style, note, rowSpan = 1, colSpan = 1 } = props;

        // Create a template for all properties to be applied
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
            validation: null,
            border: style.props.border,
            numberFormat: null
        };

        if (type.component === Checkbox) {
            cellData.validation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
        } else if (type.component === Dropdown) {
            cellData.validation = SpreadsheetApp.newDataValidation().requireValueInList(type.props.values).build();
        } else if (type.component === DatePicker) {
            cellData.validation = SpreadsheetApp.newDataValidation().requireDate().build();
            if (type.props.format) {
                cellData.numberFormat = type.props.format;
            }
        }

        // Iterate over the spanned area and apply properties to all cells in the span
        for (let rOffset = 0; rOffset < rowSpan; rOffset++) {
            for (let cOffset = 0; cOffset < colSpan; cOffset++) {
                const r = (cell.row - minRow) + rOffset;
                const c = (cell.col - minCol) + cOffset;

                // The first cell in a span gets the value, the rest are blank
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
                numberFormats[r][c] = cellData.numberFormat;
                
                // For borders, we build a boolean grid for the setBorder call
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
            merges.push(targetSheet.getRange(cell.row, cell.col, rowSpan, colSpan));
        }
    });

    // 4. Make Batched API Calls
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
        .setNumberFormats(numberFormats);

    // Sanitize the wrap strategies grid and call the API.
    const wrapStrategyMap = {
        'WRAP': SpreadsheetApp.WrapStrategy.WRAP,
        'OVERFLOW': SpreadsheetApp.WrapStrategy.OVERFLOW,
        'CLIP': SpreadsheetApp.WrapStrategy.CLIP
    };

    const sanitizedWrapStrategies = wrapStrategies.map(row => 
        row.map(strategy => wrapStrategyMap[strategy] || SpreadsheetApp.WrapStrategy.OVERFLOW)
    );
    fullRange.setWrapStrategies(sanitizedWrapStrategies);

    // Process borders separately
    for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
            const borderInfo = borders[r][c];
            if (borderInfo) {
                const range = targetSheet.getRange(minRow + r, minCol + c);
                const { top, bottom, left, right } = borderInfo;
                range.setBorder(
                    top ? true : null,
                    left ? true : null,
                    bottom ? true : null,
                    right ? true : null,
                    false, // vertical (within a range)
                    false, // horizontal (within a range)
                    top ? top.color : (right ? right.color : (bottom ? bottom.color : (left ? left.color : null))),
                    top ? top.style : (right ? right.style : (bottom ? bottom.style : (left ? left.style : null)))
                );
            }
        }
    }

    // 5. Handle Merges
    merges.forEach(range => range.merge());
}

/**
 * The main render function for the framework.
 * @param {object} rootComponent - The top-level layout component for the entire sheet.
 * @param {Sheet} targetSheet - The Google Apps Script sheet object to render to.
 */
function render(rootComponent, targetSheet) {
    if (!rootComponent || rootComponent.category !== Category.LAYOUT) {
        throw new Error("Render function requires a root Layout component.");
    }
    if (!targetSheet) {
        throw new Error("Render function requires a target sheet object.");
    }
    
    const resolvedCells = layoutEngine(rootComponent);
    commitEngine(resolvedCells, targetSheet);
}
