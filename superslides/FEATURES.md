# SuperSlides — Feature Documentation

SuperSlides is a Google Slides add-on that provides precise layout tools for aligning, sizing, and positioning shapes. This document details each implemented feature.

---

## Table of Contents

1. [Alignment Features](#alignment-features)
   - [Align Left](#align-left)
   - [Align Right](#align-right)
   - [Align Center](#align-center)
   - [Align Top](#align-top)
   - [Align Bottom](#align-bottom)
   - [Align Middle](#align-middle)
2. [Table-Based Alignment](#table-based-alignment)
   - [Align in Table](#align-in-table)
3. [Matrix Alignment](#matrix-alignment)
4. [Anchor Management](#anchor-management)
   - [Set Anchor](#set-anchor)
   - [Clear Anchor](#clear-anchor)
   - [Anchor Status](#anchor-status)
5. [How Anchor Resolution Works](#how-anchor-resolution-works)
6. [API Reference](#api-reference)
7. [Known Limitations](#known-limitations)

---

## Alignment Features

Alignment features move selected shapes so their edges line up with a reference shape (the "anchor"). All alignment operations require **at least 2 shapes selected** — one acts as the anchor, and the others move to align with it.

### Align Left

**What it does:**  
Moves all selected shapes so their **left edges** align with the anchor's left edge.

**Visual example:**

```
BEFORE:                         AFTER:

[Anchor    ]                    [Anchor    ]
      [Shape 1]          →      [Shape 1]
   [Shape 2   ]                 [Shape 2   ]

All left edges now line up vertically.
```

**How to use:**

1. Select 2 or more shapes on your slide (Shift+click or drag to select)
2. *(Optional)* Click **Set Anchor** to designate a specific shape as the reference
3. Click **Align Left** in the sidebar
4. All non-anchor shapes move horizontally so their left edges match the anchor's left edge

**Technical details:**

- The anchor's left position is retrieved using `anchor.getLeft()`
- Each other shape's left position is set to match: `el.setLeft(anchorLeft)`
- Shapes move **horizontally only** — their vertical position (top) stays the same
- The anchor shape itself does not move

---

### Align Right

**What it does:**  
Moves all selected shapes so their **right edges** align with the anchor's right edge.

**Visual example:**

```
BEFORE:                         AFTER:

    [Anchor    ]                    [Anchor    ]
[Shape 1]                →              [Shape 1]
  [Shape 2   ]                      [Shape 2   ]

All right edges now line up vertically.
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate a specific shape as the reference
3. Click **Align Right** in the sidebar
4. All non-anchor shapes move horizontally so their right edges match the anchor's right edge

**Technical details:**

- The anchor's right edge is calculated: `anchorRight = anchor.getLeft() + anchor.getWidth()`
- Each shape's new left position is calculated to place its right edge at the target:  
  `newLeft = anchorRight - el.getWidth()`
- Shapes move **horizontally only** — their vertical position stays the same
- Shapes of different widths will have different left positions, but identical right edges

---

### Align Center

**What it does:**  
Moves all selected shapes so their **horizontal centers** align with the anchor's horizontal center.

**Visual example:**

```
BEFORE:                         AFTER:

    [Anchor]                        [Anchor]
[Shape1]            →               [Shape1]
  [Shape2]                          [Shape2]

All horizontal centers now line up vertically (same X-axis).
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate a specific shape as the reference
3. Click **Align Center** in the sidebar
4. All non-anchor shapes move horizontally so their centers align with the anchor's center

**Technical details:**

- The anchor's horizontal center is calculated: `anchorCenterX = anchor.getLeft() + (anchor.getWidth() / 2)`
- Each shape's new left position is calculated to center it:  
  `newLeft = anchorCenterX - (el.getWidth() / 2)`
- Shapes move **horizontally only** — their vertical position stays the same
- Shapes of different widths will have different left/right positions, but identical center points

**Note:** This is center-based alignment, not edge-based. The center point of each shape aligns, regardless of shape width.

---

### Align Top

**What it does:**  
Moves all selected shapes so their **top edges** align with the anchor's top edge.

**Visual example:**

```
BEFORE:                              AFTER:

[Anchor]                             [Anchor]  [Shape 1]  [Shape 2]
            [Shape 1]         →      
                                     All top edges now line up horizontally.
        [Shape 2]
```

**How to use:**

1. Select 2 or more shapes on your slide (Shift+click or drag to select)
2. *(Optional)* Click **Set Anchor** to designate a specific shape as the reference
3. Click **Align Top** in the sidebar
4. All non-anchor shapes move vertically so their top edges match the anchor's top edge

**Technical details:**

- The anchor's top position is retrieved using `anchor.getTop()`
- Each other shape's top position is set to match: `el.setTop(anchorTop)`
- Shapes move **vertically only** — their horizontal position (left) stays the same
- The anchor shape itself does not move

**Note:** This is the vertical equivalent of Align Left.

---

### Align Bottom

**What it does:**  
Moves all selected shapes so their **bottom edges** align with the anchor's bottom edge.

**Visual example:**

```
BEFORE:                              AFTER:

[Anchor]                             
            [Shape 1]         →      All bottom edges now line up horizontally.
                                     
        [Shape 2]                    [Anchor]  [Shape 1]  [Shape 2]
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate a specific shape as the reference
3. Click **Align Bottom** in the sidebar
4. All non-anchor shapes move vertically so their bottom edges match the anchor's bottom edge

**Technical details:**

- The anchor's bottom edge is calculated: `anchorBottom = anchor.getTop() + anchor.getHeight()`
- Each shape's new top position is calculated to place its bottom edge at the target:  
  `newTop = anchorBottom - el.getHeight()`
- Shapes move **vertically only** — their horizontal position stays the same
- Shapes of different heights will have different top positions, but identical bottom edges

**Note:** This is the vertical equivalent of Align Right.

---

### Align Middle

**What it does:**  
Moves all selected shapes so their **vertical centers** align with the anchor's vertical center.

**Visual example:**

```
BEFORE:                         AFTER:

[Anchor]                        [Anchor]
        [Shape1]         →      [Shape1]
    [Shape2]                    [Shape2]

All vertical centers now line up horizontally (same Y-axis).
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate a specific shape as the reference
3. Click **Align Middle** in the sidebar
4. All non-anchor shapes move vertically so their centers align with the anchor's center

**Technical details:**

- The anchor's vertical center is calculated: `anchorCenterY = anchor.getTop() + (anchor.getHeight() / 2)`
- Each shape's new top position is calculated to center it:  
  `newTop = anchorCenterY - (el.getHeight() / 2)`
- Shapes move **vertically only** — their horizontal position stays the same
- Shapes of different heights will have different top/bottom positions, but identical center points

**Note:** This is center-based alignment, not edge-based. The center point of each shape aligns, regardless of shape height.

---

## Table-Based Alignment

Table-based alignment features align shapes within their containing table cells. These features are useful when you have shapes placed inside table cells and want to align them relative to the cell boundaries.

**How it works:**
- SuperSlides detects which table cell contains each shape (by checking if the shape's center point is within the cell)
- Shapes are grouped by their containing cell
- Each group is aligned within its cell with configurable padding
- Multiple shapes in the same cell are stacked vertically and **centered vertically** within the cell
- **Important:** Select only the shapes you want to align, NOT the table itself

### Align in Table

**What it does:**  
Aligns selected shapes within their containing table cells with a choice of left, center, or right alignment and configurable padding. Shapes are automatically centered vertically within their cells.

**Visual examples:**

**Left alignment (padding=10pt):**
```
┌─────────────┐
│[Shape]      │  ← 10pt from left, centered vertically
│             │
└─────────────┘
```

**Center alignment:**
```
┌─────────────┐
│   [Shape]   │  ← Centered horizontally and vertically
│             │
└─────────────┘
```

**Right alignment (padding=10pt):**
```
┌─────────────┐
│      [Shape]│  ← 10pt from right, centered vertically
│             │
└─────────────┘
```

**How to use:**

1. Create a table on your slide
2. Place shapes inside table cells (shapes don't need to be perfectly positioned)
3. Select **only the shapes** you want to align (do NOT select the table)
4. Click **Align in Table** button in the sidebar
5. A dialog will appear with options:
   - **Alignment:** Choose Left, Center, or Right (radio buttons)
   - **Padding:** Enter padding in points (default: 10)
6. Click **Apply**

**Technical details:**

- Uses shape center point to determine which cell contains it
- Tables are automatically filtered out from selection (prevents accidental table movement)
- Shapes not contained in any cell are skipped
- Multiple shapes in the same cell are stacked vertically and centered as a group
- Vertical centering: `currentTop = cellTop + (cellHeight / 2) - (totalHeight / 2)`
- Horizontal alignment:
  - **Left:** `cellLeft + padding`
  - **Right:** `cellLeft + cellWidth - shapeWidth - padding`
  - **Center:** `cellLeft + (cellWidth / 2) - (shapeWidth / 2)`
- Table dimensions calculated by summing individual column widths and row heights (handles non-uniform cells)

---

## Matrix Alignment

**What it does:**  
Arranges selected shapes into a grid/matrix layout with user-specified dimensions and spacing.

**Visual example:**

```
BEFORE:                         AFTER (3×3 matrix):

[Shape1] [Shape2] [Shape3]     [Shape1] [Shape2] [Shape3]
[Shape4] [Shape5] [Shape6] →   [Shape4] [Shape5] [Shape6]
[Shape7] [Shape8] [Shape9]     [Shape7] [Shape8] [Shape9]
```

**How to use:**

1. Select the shapes you want to arrange (maintains selection order)
2. Specify desired rows and columns (e.g., 3 rows, 3 columns)
3. Specify spacing between shapes (default: 20 points)
4. Click **Arrange in Matrix**

**Auto-expansion:**

If you select more shapes than fit in the requested dimensions, rows are automatically expanded:

- **Example:** 10 shapes, 3×3 requested → becomes 4×3 (4 rows, 3 columns)
- Last row may be incomplete (e.g., 10 shapes in 3 cols = 4 rows, last row has 1 shape)
- Formula: `actualRows = Math.ceil(numShapes / requestedCols)`

**Technical details:**

- Shapes maintain their **selection order** (not sorted by position)
- Matrix is positioned using the **bounding box** of current selection
- Cell size is calculated: `cellWidth = (bounds.width - (cols-1)*spacing) / cols`
- Position calculation: `left = bounds.left + col*(cellWidth + spacing)`

**Use cases:**

- Creating uniform grids of icons or images
- Organizing scattered shapes into a neat layout
- Creating dashboard-style arrangements

---

## Anchor Management

The "anchor" is the reference shape that other shapes align to. SuperSlides provides explicit anchor control so you always know which shape is the reference.

### Set Anchor

**What it does:**  
Saves the currently selected shape as the anchor. This anchor persists until you clear it or set a new one.

**How to use:**

1. Select a single shape (the one you want as your reference)
2. Click **Set Anchor** in the sidebar
3. The status will show "Anchor is set."

**Technical details:**

- Stores the shape's unique object ID in `DocumentProperties`
- If multiple shapes are selected, the **first** element in the selection array is used
- The anchor ID persists per presentation (survives closing/reopening the sidebar)
- All users editing the same presentation share the same anchor

---

### Clear Anchor

**What it does:**  
Removes the stored anchor. Future alignment operations will use the fallback behavior (last element in selection becomes the anchor).

**How to use:**

1. Click **Clear Anchor** in the sidebar
2. The status will show "No anchor set."

**Technical details:**

- Deletes the `anchorId` key from `DocumentProperties`
- Does not affect any shapes — only removes the stored reference

---

### Anchor Status

**What it does:**  
Displays whether an anchor is currently set.

**How it appears:**

- **"Anchor is set."** — An anchor has been explicitly set
- **"No anchor set."** — No anchor; operations will use fallback behavior

**Technical details:**

- Status is fetched from the server when the sidebar loads
- Status updates automatically after Set Anchor or Clear Anchor operations

---

## How Anchor Resolution Works

When you perform an alignment operation, SuperSlides needs to determine which shape is the anchor. This is handled by the `getAnchorOrFallback()` function.

### Resolution Rules

```
┌─────────────────────────────────────────────────────────────┐
│                    Anchor Resolution                         │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1. Is there an explicit anchor set?                        │
│     │                                                       │
│     ├─ YES → Is that anchor in the current selection?       │
│     │        │                                              │
│     │        ├─ YES → Use the explicit anchor ✓             │
│     │        │                                              │
│     │        └─ NO  → Fall through to fallback              │
│     │                                                       │
│     └─ NO  → Use fallback                                   │
│                                                             │
│  2. Fallback: Use the LAST element in the selection array   │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### Why This Design?

- **Explicit anchor takes priority**: When you set an anchor, you expect it to be used
- **Anchor must be in selection**: Prevents confusion when anchor is on a different slide
- **Fallback exists for quick testing**: You don't *have* to set an anchor for basic use
- **Fallback is imperfect by design**: Google doesn't guarantee selection order, so "last element" may not be what you expect. This encourages using explicit anchors for precision.

---

## API Reference

### Server-Side Functions (Code.gs)

| Function | Parameters | Returns | Description |
|----------|------------|---------|-------------|
| `onOpen()` | — | — | Creates the SuperSlides menu on document open |
| `showSidebar()` | — | — | Displays the sidebar UI |
| `setAnchor()` | — | `string` | Stores selected element as anchor |
| `clearAnchor()` | — | `string` | Removes stored anchor |
| `getAnchorStatus()` | — | `string` | Returns current anchor state |
| `getAnchorOrFallback(elements)` | `PageElement[]` | `PageElement` | Resolves which element is the anchor |
| `alignLeft()` | — | `string` | Aligns selection to anchor's left edge |
| `alignRight()` | — | `string` | Aligns selection to anchor's right edge |
| `alignCenter()` | — | `string` | Aligns selection to anchor's horizontal center |
| `alignTop()` | — | `string` | Aligns selection to anchor's top edge |
| `alignBottom()` | — | `string` | Aligns selection to anchor's bottom edge |
| `alignMiddle()` | — | `string` | Aligns selection to anchor's vertical center |

### Client-Side Functions (Sidebar.html)

| Function | Description |
|----------|-------------|
| `refreshStatus()` | Fetches and displays anchor status |
| `runSetAnchor()` | Calls `setAnchor()`, then refreshes status |
| `runClearAnchor()` | Calls `clearAnchor()`, then refreshes status |
| `runAlignLeft()` | Shows "Aligning...", calls `alignLeft()`, shows result |
| `runAlignRight()` | Shows "Aligning...", calls `alignRight()`, shows result |
| `runAlignCenter()` | Shows "Aligning...", calls `alignCenter()`, shows result |
| `runAlignTop()` | Shows "Aligning...", calls `alignTop()`, shows result |
| `runAlignBottom()` | Shows "Aligning...", calls `alignBottom()`, shows result |
| `runAlignMiddle()` | Shows "Aligning...", calls `alignMiddle()`, shows result |
| `showTableAlignDialog()` | Shows modal dialog for table alignment options |
| `hideTableAlignDialog()` | Hides the table alignment modal dialog |
| `runTableAlign()` | Gets selected alignment and padding, calls `alignToTable()` |
| `runArrangeMatrix()` | Gets rows/cols/spacing, calls `arrangeMatrix()`, shows result |

---

## Known Limitations

### Selection Order is Unpredictable

Google Apps Script does not return shapes in the order they were clicked. When you shift-click shapes A, B, C, the array might be `[B, C, A]` or any other order. This is why:

- The fallback anchor (last element) may not be what you expect
- **Recommendation:** Always use **Set Anchor** for predictable results

### Anchor Must Be in Current Selection

If you set an anchor on Slide 1, then try to align shapes on Slide 2, the anchor won't be used (it's not in the current selection). The operation will fall back to using the last element.

### Minimum 2 Elements Required

Alignment operations require at least 2 selected shapes:
- 1 shape = nothing to align to
- 2+ shapes = anchor + elements to move

### Some Elements May Not Move

Certain element types (like embedded videos or linked objects) may not support `setLeft()` or `setTop()`. These are caught by try/catch and reported in the status message.

### Table Alignment Limitations

- **Shape must be in a cell**: Only shapes whose center point is within a table cell are aligned. Shapes outside cells are skipped.
- **Cell detection uses center point**: If a shape spans multiple cells, the cell containing its center point is used.
- **Multiple tables**: If multiple tables exist on a slide, all are checked, but shapes align only within their containing table.
- **Select shapes, not tables**: The table itself should not be selected. Tables are automatically filtered from the selection to prevent accidental movement.
- **Vertical centering**: Shapes are always centered vertically within their cells. Top/bottom alignment within cells is not currently supported.

### Matrix Alignment Limitations

- **Selection order matters**: Shapes are arranged in the order they were selected, not sorted by position.
- **Matrix bounds**: The matrix uses the bounding box of current selection as its reference area.
- **Incomplete rows**: If shapes don't perfectly fit the dimensions, the last row may be incomplete (expected behavior).

---

## File Structure

```
superslides/
├── src/
│   ├── appsscript.json    # Manifest (scopes, runtime)
│   ├── Code.gs            # Server-side logic (~1450 lines)
│   └── Sidebar.html       # UI + client-side JS (~570 lines)
└── FEATURES.md            # This documentation
```

---

*Last updated: January 2026*
