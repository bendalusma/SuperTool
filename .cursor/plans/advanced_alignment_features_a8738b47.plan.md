---
name: Advanced Alignment Features
overview: "Implement two advanced alignment features: table-based alignment (align shapes within their containing table cells with padding) and matrix alignment (arrange shapes into a grid with user-specified dimensions and spacing)."
todos:
  - id: table-helpers
    content: Create helper functions for table detection and cell bounds calculation
    status: completed
  - id: table-main
    content: Implement alignToTable() function with left/right/center options
    status: completed
    dependencies:
      - table-helpers
  - id: table-ui
    content: Add table alignment UI (buttons + padding input) to Sidebar.html
    status: completed
    dependencies:
      - table-main
  - id: matrix-helpers
    content: Create helper functions for matrix dimension calculation and selection bounds
    status: completed
  - id: matrix-main
    content: Implement arrangeMatrix() function with auto-expand logic
    status: completed
    dependencies:
      - matrix-helpers
  - id: matrix-ui
    content: Add matrix alignment UI (inputs + button) to Sidebar.html
    status: completed
    dependencies:
      - matrix-main
  - id: update-docs
    content: Update FEATURES.md with both new features and examples
    status: completed
    dependencies:
      - table-ui
      - matrix-ui
---

# Advanced Alignment Features Implementation Plan

## Overview

Implement two advanced alignment features that extend beyond simple edge/center alignment:

1. **Table-Based Alignment** - Align shapes within their containing table cells
2. **Matrix Alignment** - Arrange shapes into a grid/matrix layout

---

## Feature 1: Table-Based Alignment

### Requirements

- Detect which table cell contains each shape (using shape center point)
- Align shapes within their cells (left/right/center) with configurable padding
- Skip shapes not contained in any table cell
- Handle multiple shapes in the same cell (stack them)

### Implementation Steps

#### 1.1 Helper Functions (`Code.gs`)

- `findContainingCell(shape, tables)` - Finds which cell contains a shape's center point
- Iterate through all tables on the slide
- For each table, check each cell's bounds
- Return cell object if center point is within cell bounds
- Return null if not found
- `getCellBounds(cell)` - Extract left, top, width, height of a cell
- `alignShapesInCell(shapes, cell, alignment, padding)` - Align multiple shapes within a cell

#### 1.2 Main Function (`Code.gs`)

- `alignToTable(alignment, padding)` - Main function
- Get selected shapes
- Get all tables on current slide
- Group shapes by containing cell
- For each cell group, align shapes with specified padding
- Return status message

#### 1.3 UI (`Sidebar.html`)

- Add section "Table Alignment"
- Three buttons: "Align Left in Table", "Align Right in Table", "Align Center in Table"
- Input field for padding (default: 10pt)
- Handler functions: `runAlignToTableLeft()`, `runAlignToTableRight()`, `runAlignToTableCenter()`

---

## Feature 2: Matrix Alignment

### Requirements

- User specifies rows and columns
- Auto-expand dimensions if shapes don't fit (e.g., 10 shapes in 3x3 → 4x3)
- Maintain selection order
- Configurable spacing between shapes
- Position matrix using selection bounds as reference

### Implementation Steps

#### 2.1 Helper Functions (`Code.gs`)

- `calculateMatrixDimensions(numShapes, requestedRows, requestedCols)` - Calculate actual dimensions
- If `numShapes > requestedRows * requestedCols`, auto-expand
- Calculate: `actualRows = Math.ceil(numShapes / requestedCols)`
- Return `{rows, cols}`
- `getSelectionBounds(elements)` - Get bounding box of all selected shapes
- Find min(left), min(top), max(right), max(bottom)
- Return `{left, top, width, height}`
- `arrangeInMatrix(elements, rows, cols, spacing, bounds)` - Arrange shapes in grid
- Calculate cell size: `cellWidth = (bounds.width - (cols-1)*spacing) / cols`
- Calculate cell size: `cellHeight = (bounds.height - (rows-1)*spacing) / rows`
- For each shape, calculate its position: `row = Math.floor(index / cols)`, `col = index % cols`
- Position shape: `left = bounds.left + col*(cellWidth + spacing)`, `top = bounds.top + row*(cellHeight + spacing)`

#### 2.2 Main Function (`Code.gs`)

- `arrangeMatrix(rows, cols, spacing)` - Main function
- Get selected shapes
- Validate inputs (rows > 0, cols > 0, spacing >= 0)
- Calculate actual dimensions (with auto-expand)
- Get selection bounds
- Arrange shapes in matrix
- Return status message

#### 2.3 UI (`Sidebar.html`)

- Add section "Matrix Alignment"
- Input fields: Rows, Columns, Spacing (default: 20pt)
- Button: "Arrange in Matrix"
- Handler function: `runArrangeMatrix()`
- Show preview/confirmation if dimensions need to be adjusted

---

## File Changes

### `superslides/src/Code.gs`

- Add helper functions for table detection and cell bounds
- Add `alignToTable()` function
- Add helper functions for matrix calculation
- Add `arrangeMatrix()` function
- Maintain detailed comments matching existing style

### `superslides/src/Sidebar.html`

- Add "Table Alignment" section with 3 buttons + padding input
- Add "Matrix Alignment" section with inputs + button
- Add client-side handler functions
- Add input validation and error handling

### `superslides/FEATURES.md`

- Document both new features
- Add visual examples
- Update API reference table
- Add usage instructions

---

## Technical Considerations

### Table Detection

- Google Slides API: `Slide.getTables()` returns all tables
- `Table.getCell(rowIndex, columnIndex)` to access cells
- `TableCell.getContent()` returns content, but we need bounds
- Challenge: Cell bounds not directly available - need to calculate from table position + cell dimensions
- Solution: Use `Table.getLeft()`, `Table.getTop()`, `Table.getWidth()`, `Table.getHeight()` and calculate cell positions

### Matrix Overflow Handling

- Example: 10 shapes, 3x3 requested → 4 rows, 3 cols (last row has 1 shape)
- Formula: `actualRows = Math.ceil(numShapes / cols)`
- Last row may be incomplete - that's expected behavior

### Error Handling

- No tables found → Show error
- No shapes in cells → Show warning
- Invalid matrix dimensions → Validate and show error
- Negative spacing → Set to 0 or show error

---

## Testing Scenarios

### Table Alignment

1. Single shape in cell → Aligns correctly
2. Multiple shapes in same cell → Stacked correctly
3. Shape not in any cell → Skipped
4. Multiple tables on slide → Checks all tables

### Matrix Alignment

1. Perfect fit (9 shapes, 3x3) → Arranges correctly
2. Overflow (10 shapes, 3x3) → Auto-expands to 4x3
3. Underflow (6 shapes, 3x3) → Uses 3x2, leaves empty cells
4. Single row/column → Works correctly

---

## Implementation Order

1. **Table Alignment** (simpler, fewer edge cases)

- Helper functions
- Main function
- UI integration
- Testing

2. **Matrix Alignment** (more complex, needs dimension calculation)

- Helper functions
- Main function
- UI integration
- Testing

3. **Documentation**

- Update FEATURES.md
- Add examples
- Update API reference