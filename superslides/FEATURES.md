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
2. [Position Manipulation](#position-manipulation)
   - [Swap Positions](#swap-positions)
3. [Distribution](#distribution)
   - [Distribute Horizontally](#distribute-horizontally)
   - [Distribute Vertically](#distribute-vertically)
4. [Docking](#docking)
   - [Dock Left](#dock-left)
   - [Dock Right](#dock-right)
   - [Dock Top](#dock-top)
   - [Dock Bottom](#dock-bottom)
5. [Table-Based Alignment](#table-based-alignment)
   - [Align in Table](#align-in-table)
   - [Swap Rows in Table](#swap-rows-in-table)
   - [Swap Columns in Table](#swap-columns-in-table)
6. [Matrix Alignment](#matrix-alignment)
7. [Slide Layout](#slide-layout)
   - [Insert 2 Columns](#insert-2-columns)
   - [Insert 3 Columns](#insert-3-columns)
   - [Insert 4 Columns](#insert-4-columns)
   - [Insert Footnote](#insert-footnote)
8. [Size Manipulation](#size-manipulation)
   - [Align Width](#align-width)
   - [Align Height](#align-height)
   - [Align Both](#align-both)
9. [Stretch Objects](#stretch-objects)
   - [Stretch Left](#stretch-left)
   - [Stretch Right](#stretch-right)
   - [Stretch Top](#stretch-top)
   - [Stretch Bottom](#stretch-bottom)
10. [Fill Space](#fill-space)
   - [Fill Left](#fill-left)
   - [Fill Right](#fill-right)
   - [Fill Top](#fill-top)
   - [Fill Bottom](#fill-bottom)
11. [Magic Resizer](#magic-resizer)
12. [Anchor Management](#anchor-management)
   - [Set Anchor](#set-anchor)
   - [Clear Anchor](#clear-anchor)
   - [Anchor Status](#anchor-status)
13. [How Anchor Resolution Works](#how-anchor-resolution-works)
14. [API Reference](#api-reference)
15. [Known Limitations](#known-limitations)

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

## Position Manipulation

Position manipulation features allow you to move and rearrange shapes in ways that go beyond simple alignment.

### Swap Positions

**What it does:**  
Swaps the positions (left, top coordinates) of exactly two selected shapes. Size, rotation, and styling remain unchanged.

**Visual example:**

```
BEFORE:                         AFTER:

[Shape A]                       [Shape B]
              [Shape B]   →                   [Shape A]

Shapes exchange positions while maintaining their size and other properties.
```

**How to use:**

1. Select exactly 2 shapes on your slide (Shift+click to select both)
2. Click **Swap Positions** in the sidebar
3. The shapes will exchange their left/top coordinates instantly

**Technical details:**

- Stores both shapes' positions: `(left1, top1)` and `(left2, top2)`
- Applies swap: `shape1.setLeft(left2)`, `shape1.setTop(top2)`, and vice versa
- Only position is swapped - size, rotation, fill color, borders, and all other properties remain unchanged
- Shapes may overlap after swapping if they have different sizes (this is expected behavior)

**Use cases:**
- Quick position exchange without manual dragging
- Swapping placeholder positions in layouts (e.g., swap header and footer)
- A/B testing different visual arrangements
- Reversing the order of two elements in a design

**Error handling:**
- If 0 or 1 shape is selected: Shows error "Please select exactly 2 shapes"
- If more than 2 shapes are selected: Shows error with current count

---

## Distribution

Distribution features arrange selected shapes with equal spacing between them, similar to PowerPoint's built-in distribute functionality. These features are essential for creating professional, evenly-spaced layouts.

### Distribute Horizontally

**What it does:**  
Arranges selected shapes so they have equal horizontal gaps between them. The leftmost and rightmost shapes stay fixed as anchors; middle shapes are repositioned to create even spacing.

**Visual example:**

```
BEFORE:                         AFTER:

[1]  [2][3]    [4]       [5]   [1]  [2]  [3]  [4]  [5]
                          →     
Random, uneven spacing          Equal gaps between all shapes
```

**How to use:**

1. Select at least 3 shapes on your slide (Shift+click or drag to select)
2. Click **Distribute Horizontally** in the sidebar
3. Shapes will be repositioned with equal gaps

**Technical details:**

- Requires minimum 3 shapes (2 anchors + at least 1 to distribute)
- Shapes are sorted left-to-right automatically based on their left position
- Formula:
  ```
  totalSpace = rightmost.right - leftmost.left
  totalShapeWidth = sum of all widths
  gapSize = (totalSpace - totalShapeWidth) / (count - 1)
  ```
- Each shape is positioned at: `previousRight + gapSize`
- Vertical positions remain unchanged (shapes don't move up/down)

**Edge cases:**
- If shapes don't fit in available space (total width > total space), they will overlap (same as PowerPoint behavior)
- Single pixel rounding may occur on very small gaps due to floating point precision
- Works with shapes of different widths - gaps are equal, not positions

**Use cases:**
- Creating evenly spaced button rows in navigation bars
- Aligning timeline markers or milestones
- Organizing icons in a toolbar
- Spacing out steps in a process diagram

---

### Distribute Vertically

**What it does:**  
Arranges selected shapes so they have equal vertical gaps between them. The topmost and bottommost shapes stay fixed as anchors; middle shapes are repositioned to create even spacing.

**Visual example:**

```
BEFORE:                AFTER:

[Shape 1]              [Shape 1]
[Shape 2]                ↓ equal gap
                       [Shape 2]
[Shape 3]       →        ↓ equal gap
       [Shape 4]       [Shape 3]
[Shape 5]                ↓ equal gap
                       [Shape 4]
                         ↓ equal gap
                       [Shape 5]
```

**How to use:**

1. Select at least 3 shapes on your slide (Shift+click or drag to select)
2. Click **Distribute Vertically** in the sidebar
3. Shapes will be repositioned with equal gaps

**Technical details:**

- Requires minimum 3 shapes (2 anchors + at least 1 to distribute)
- Shapes are sorted top-to-bottom automatically based on their top position
- Formula:
  ```
  totalSpace = bottommost.bottom - topmost.top
  totalShapeHeight = sum of all heights
  gapSize = (totalSpace - totalShapeHeight) / (count - 1)
  ```
- Each shape is positioned at: `previousBottom + gapSize`
- Horizontal positions remain unchanged (shapes don't move left/right)

**Edge cases:**
- If shapes don't fit in available space (total height > total space), they will overlap
- Works identically to horizontal distribution, but operates on Y-axis
- Works with shapes of different heights - gaps are equal, not positions

**Use cases:**
- Creating evenly spaced vertical navigation menus
- Aligning list items or bullet points
- Organizing vertical timelines or progress trackers
- Spacing out form fields in a vertical layout

---

## Docking

Docking features move selected shapes until they "touch" the anchor shape's edges with zero gap. These operations create precisely aligned layouts where shapes abut each other perfectly, which is essential for flowcharts, connected diagrams, and tight layouts.

**Important:** All docking operations require an anchor. Use **Set Anchor** first for predictable results. If no anchor is explicitly set, the fallback anchor (last selected shape) will be used.

### Dock Left

**What it does:**  
Moves selected shapes horizontally until their **right edges** touch the anchor's **left edge**. Creates a "stacked to the left" effect where shapes align perfectly to the left side of the anchor with zero gap.

**Visual example:**

```
BEFORE:                         AFTER:

     [Shape 1]                 [Shape 1][Anchor]
[Anchor]                  →    [Shape 2][Anchor]
   [Shape 2]

Shapes dock to the left side of the anchor, right edges touching.
```

**How to use:**

1. Select anchor shape and click **Set Anchor** (recommended for precision)
2. Select anchor + shapes to dock (Shift+click to add to selection)
3. Click **Dock Left** in the sidebar
4. Non-anchor shapes move horizontally to touch anchor's left edge

**Technical details:**

- Formula: `newLeft = anchorLeft - shapeWidth`
- Only horizontal position changes; vertical position (top) remains unchanged
- Multiple shapes will stack on top of each other at the same horizontal position (overlap vertically)
- To avoid vertical overlap, dock shapes one at a time or use Align Middle after docking

**Use cases:**
- Creating side-by-side layouts without gaps
- Aligning labels to the left of content boxes in forms
- Building flowcharts with connected shapes (shape → anchor flows)
- Creating compact horizontal sequences

**Example workflow:**
```
1. Have 3 shapes: [Label], [Box], [Arrow]
2. Set [Box] as anchor
3. Select [Label] + [Box], click Dock Left → [Label] touches left side of [Box]
4. Select [Arrow] + [Box], click Dock Right → [Arrow] touches right side of [Box]
Result: [Label][Box][Arrow] with no gaps
```

---

### Dock Right

**What it does:**  
Moves selected shapes horizontally until their **left edges** touch the anchor's **right edge**. Creates a "stacked to the right" effect where shapes align perfectly to the right side of the anchor with zero gap.

**Visual example:**

```
BEFORE:                         AFTER:

[Anchor]                       [Anchor][Shape 1]
     [Shape 1]            →    [Anchor][Shape 2]
[Shape 2]

Shapes dock to the right side of the anchor, left edges touching.
```

**How to use:**

1. Select anchor shape and click **Set Anchor** (recommended for precision)
2. Select anchor + shapes to dock (Shift+click to add to selection)
3. Click **Dock Right** in the sidebar
4. Non-anchor shapes move horizontally to touch anchor's right edge

**Technical details:**

- Formula: `newLeft = anchorLeft + anchorWidth`
- Only horizontal position changes; vertical position remains unchanged
- Multiple shapes will stack on top of each other at the same horizontal position (overlap vertically)

**Use cases:**
- Creating horizontal flowcharts with left-to-right flow
- Positioning icons or labels to the right of content
- Building timeline sequences with connected elements
- Creating expansion arrows (e.g., box → arrow → box)

---

### Dock Top

**What it does:**  
Moves selected shapes vertically until their **bottom edges** touch the anchor's **top edge**. Creates a "stacked above" effect where shapes align perfectly above the anchor with zero gap.

**Visual example:**

```
BEFORE:                         AFTER:

[Shape 1]                       [Shape 1]
                                [Shape 2]
[Anchor]                  →     [Anchor]
                  [Shape 2]
                  
Shapes dock above the anchor, bottom edges touching top.
```

**How to use:**

1. Select anchor shape and click **Set Anchor** (recommended for precision)
2. Select anchor + shapes to dock (Shift+click to add to selection)
3. Click **Dock Top** in the sidebar
4. Non-anchor shapes move vertically to touch anchor's top edge

**Technical details:**

- Formula: `newTop = anchorTop - shapeHeight`
- Only vertical position changes; horizontal position (left) remains unchanged
- Multiple shapes will stack on top of each other at the same vertical position (overlap horizontally)

**Use cases:**
- Creating stacked vertical layouts without gaps
- Building vertical flowcharts with top-to-bottom flow
- Positioning headers or titles directly above content sections
- Creating compact vertical sequences (e.g., stacked process steps)

---

### Dock Bottom

**What it does:**  
Moves selected shapes vertically until their **top edges** touch the anchor's **bottom edge**. Creates a "stacked below" effect where shapes align perfectly below the anchor with zero gap.

**Visual example:**

```
BEFORE:                         AFTER:

[Shape 1]                       [Anchor]
                                [Shape 1]
[Anchor]                  →     [Shape 2]
     [Shape 2]

Shapes dock below the anchor, top edges touching bottom.
```

**How to use:**

1. Select anchor shape and click **Set Anchor** (recommended for precision)
2. Select anchor + shapes to dock (Shift+click to add to selection)
3. Click **Dock Bottom** in the sidebar
4. Non-anchor shapes move vertically to touch anchor's bottom edge

**Technical details:**

- Formula: `newTop = anchorTop + anchorHeight`
- Only vertical position changes; horizontal position remains unchanged
- Multiple shapes will stack on top of each other at the same vertical position (overlap horizontally)

**Use cases:**
- Creating dropdown-style layouts
- Positioning footers or captions directly below content sections
- Building vertical step-by-step diagrams
- Creating stacked information hierarchies (e.g., title → subtitle → content)

**Common docking workflow:**
```
Create a vertical stack with no gaps:
1. Have 4 shapes: [Header], [Content], [Footer], [Button]
2. Set [Content] as anchor
3. Select [Header] + [Content], click Dock Top → Header touches top of Content
4. Select [Footer] + [Content], click Dock Bottom → Footer touches bottom of Content
5. Select [Button] + [Footer], set [Footer] as new anchor, click Dock Bottom
Result: Perfectly stacked vertical layout with zero gaps
```

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

### Swap Rows in Table

**What it does:**  
Swaps two rows in a selected table by row index (1-based). You can swap content only, or swap content **with source text/cell formatting**.

**How to use:**

1. Select a table on the current slide (or ensure only one table exists on the slide)
2. Click **Swap Rows in Table** in the sidebar
3. Enter:
   - **First row** (1-based)
   - **Second row** (1-based)
4. *(Optional)* Enable **Keep source text style + cell formatting (fill/alignment)**
5. Click **Apply**

**Modes:**

- **Default mode (checkbox OFF):**
  - Swaps plain text values between row cells
  - Destination row keeps its existing formatting context

- **Keep formatting mode (checkbox ON):**
  - Swaps text content and formatting payload together
  - Preserves and transfers:
    - text run styles (font family, size, weight, bold/italic/underline, color, links, etc.)
    - paragraph styles (line spacing, spacing above/below, indents, paragraph alignment)
    - cell fill
    - cell content alignment

**Technical details:**

- Uses 1-based user input and converts to 0-based table indexes
- Validates bounds against actual table row count
- If both indexes are the same: operation returns early with no change
- Merged-cell swap is blocked with a clear error message

---

### Swap Columns in Table

**What it does:**  
Swaps two columns in a selected table by column index (1-based). Supports the same two modes as row swapping.

**How to use:**

1. Select a table on the current slide (or ensure only one table exists on the slide)
2. Click **Swap Columns in Table** in the sidebar
3. Enter:
   - **First column** (1-based)
   - **Second column** (1-based)
4. *(Optional)* Enable **Keep source text style + cell formatting (fill/alignment)**
5. Click **Apply**

**Modes:**

- **Default mode (checkbox OFF):**
  - Swaps plain text values between column cells
  - Destination column keeps its existing formatting context

- **Keep formatting mode (checkbox ON):**
  - Swaps text + style payload with cell
  - Preserves/transfers the same formatting set as row swap mode

**Technical details:**

- Uses 1-based user input and converts to 0-based table indexes
- Validates bounds against actual table column count
- If both indexes are the same: operation returns early with no change
- Merged-cell swap is blocked with a clear error message

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

## Slide Layout

Slide Layout features allow you to quickly insert pre-formatted text boxes to create common slide layouts. These features are designed for rapid slide creation with consistent formatting and positioning.

**Key benefits:**
- Instant layout creation with one click
- Consistent spacing and sizing across slides
- Pre-formatted text with appropriate font sizes
- Professional-looking layouts without manual positioning

**Layout specifications:**
- All layouts positioned 1.4 inches from top of slide (below slide title)
- Left margin: 0.33 inches from left edge
- Right margin: 0.33 inches from right edge
- Spread across 0.33 to 9.67 inches horizontal space (9.34" total width)
- Title boxes: 0.6 inches tall, 16pt bold text
- Content boxes: 2.95 inches tall, 14pt text (ends at 5.0" to avoid footnote overlap)
- Small gap (0.05") between title and content boxes
- Gap between columns: 0.2 inches
- Internal padding: 0.01 inches on all sides of text boxes

### Insert 2 Columns

**What it does:**  
Creates a 2-column layout with title and content text boxes. Each column contains a title box and a content box, pre-formatted and properly positioned.

**Visual example:**

```
┌──────────────────────┐  ┌──────────────────────┐
│                      │  │                      │
│     Header 1         │  │     Header 2         │  ← Title boxes (1")
│     (16pt, bold)     │  │     (16pt, bold)     │
│                      │  │                      │
├──────────────────────┤  ├──────────────────────┤
│                      │  │                      │
│                      │  │                      │
│     Content 1        │  │     Content 2        │  ← Content boxes (3")
│     (14pt)           │  │     (14pt)           │
│                      │  │                      │
│                      │  │                      │
└──────────────────────┘  └──────────────────────┘

Total: 4 text boxes (2 titles + 2 content boxes)
Column width: ~4.57 inches each
Left margin: 0.33 inches, Right margin: 0.33 inches
```

**How to use:**

1. Click on a slide in the main canvas (not master or layout)
2. Click **Insert 2 Columns** in the sidebar
3. The layout will be automatically created with placeholder text
4. Edit the placeholder text as needed

**What gets created:**
- **2 title boxes:** "Header 1" and "Header 2" (bold, 16pt)
- **2 content boxes:** "Content 1" and "Content 2" (14pt)

**Technical details:**

- Left margin: 0.33 inches from left edge
- Column width calculation: `(9.34 - 0.2) / 2 = 4.57 inches`
- Vertical start position: 1.4 inches from top (below slide title)
- Title box position: 1.4 inches from top, height: 0.6 inches
- Content box position: 1.4 + 0.6 + 0.05 = 2.05 inches from top, height: 2.95 inches
- Content ends at: 2.05 + 2.95 = 5.0 inches (leaves room for footnote at 5.05")
- Internal padding: 0.01 inches on all sides of text
- All text boxes are fully editable after creation

**Use cases:**
- Side-by-side comparisons (before/after, pros/cons)
- Two-topic presentations
- Feature comparison slides
- Dual-path process diagrams

---

### Insert 3 Columns

**What it does:**  
Creates a 3-column layout with title and content text boxes. Each column contains a title box and a content box, providing a balanced three-part layout.

**Visual example:**

```
┌──────────────┐  ┌──────────────┐  ┌──────────────┐
│              │  │              │  │              │
│  Header 1    │  │  Header 2    │  │  Header 3    │  ← Title boxes (1")
│  (16pt,bold) │  │  (16pt,bold) │  │  (16pt,bold) │
│              │  │              │  │              │
├──────────────┤  ├──────────────┤  ├──────────────┤
│              │  │              │  │              │
│              │  │              │  │              │
│  Content 1   │  │  Content 2   │  │  Content 3   │  ← Content boxes (3")
│  (14pt)      │  │  (14pt)      │  │  (14pt)      │
│              │  │              │  │              │
│              │  │              │  │              │
└──────────────┘  └──────────────┘  └──────────────┘

Total: 6 text boxes (3 titles + 3 content boxes)
Column width: ~2.98 inches each
Left margin: 0.33 inches, Right margin: 0.33 inches
```

**How to use:**

1. Click on a slide in the main canvas (not master or layout)
2. Click **Insert 3 Columns** in the sidebar
3. The layout will be automatically created with placeholder text
4. Edit the placeholder text as needed

**What gets created:**
- **3 title boxes:** "Header 1", "Header 2", "Header 3" (bold, 16pt)
- **3 content boxes:** "Content 1", "Content 2", "Content 3" (14pt)

**Technical details:**

- Left margin: 0.33 inches from left edge
- Column width calculation: `(9.34 - 2 * 0.2) / 3 = 2.98 inches`
- Gap between columns: 0.2 inches
- Title height: 0.6 inches, Content height: 2.95 inches
- Internal padding: 0.01 inches on all sides
- Same vertical positioning as 2-column layout (1.4" from top)

**Use cases:**
- Three-step processes (Plan → Execute → Review)
- Past/Present/Future comparisons
- Three-feature showcases
- Triple comparison slides (Good/Better/Best)

---

### Insert 4 Columns

**What it does:**  
Creates a 4-column layout with title and content text boxes. Each column contains a title box and a content box, ideal for multi-option comparisons or four-part processes.

**Visual example:**

```
┌─────────┐  ┌─────────┐  ┌─────────┐  ┌─────────┐
│         │  │         │  │         │  │         │
│Header 1 │  │Header 2 │  │Header 3 │  │Header 4 │  ← Title boxes (1")
│(16pt,b) │  │(16pt,b) │  │(16pt,b) │  │(16pt,b) │
│         │  │         │  │         │  │         │
├─────────┤  ├─────────┤  ├─────────┤  ├─────────┤
│         │  │         │  │         │  │         │
│         │  │         │  │         │  │         │
│Content 1│  │Content 2│  │Content 3│  │Content 4│  ← Content boxes (3")
│(14pt)   │  │(14pt)   │  │(14pt)   │  │(14pt)   │
│         │  │         │  │         │  │         │
│         │  │         │  │         │  │         │
└─────────┘  └─────────┘  └─────────┘  └─────────┘

Total: 8 text boxes (4 titles + 4 content boxes)
Column width: ~2.18 inches each
Left margin: 0.33 inches, Right margin: 0.33 inches
```

**How to use:**

1. Click on a slide in the main canvas (not master or layout)
2. Click **Insert 4 Columns** in the sidebar
3. The layout will be automatically created with placeholder text
4. Edit the placeholder text as needed

**What gets created:**
- **4 title boxes:** "Header 1" through "Header 4" (bold, 16pt)
- **4 content boxes:** "Content 1" through "Content 4" (14pt)

**Technical details:**

- Left margin: 0.33 inches from left edge
- Column width calculation: `(9.34 - 3 * 0.2) / 4 = 2.18 inches`
- Gap between columns: 0.2 inches
- Title height: 0.6 inches, Content height: 2.95 inches
- Internal padding: 0.01 inches on all sides
- Same vertical positioning as other column layouts (1.4" from top)

**Use cases:**
- Four-step processes or workflows
- Quarterly comparisons (Q1, Q2, Q3, Q4)
- Four-feature showcases
- Multi-option comparisons (Free/Basic/Pro/Enterprise)

**Note:** With 4 columns, each column is narrower (~2.125 inches). Consider using:
- Shorter text or abbreviations in headers
- Bullet points instead of paragraphs in content boxes
- Smaller font sizes if needed for readability

---

### Insert Footnote

**What it does:**  
Inserts a pre-formatted footnote text box at the bottom of the slide. Perfect for adding citations, references, disclaimers, or additional context.

**Visual example:**

```
┌─────────────────────────────────────────────────┐
│                                                 │
│         [Main slide content above]              │
│                                                 │
│                                                 │
├─────────────────────────────────────────────────┤
│ Footnotes: 1. Footnotes use 9-point text 2.    │  ← Footnote box
│ Items should be numbered... (9pt)              │     (0.5" tall)
└─────────────────────────────────────────────────┘

Position: 5.05" from top, 0.33" from left
Width: 9.34 inches
Default text: "Footnotes: 1. Footnotes use 9-point text 2. Items should be numbered and presented without commas."
```

**How to use:**

1. Click on a slide in the main canvas (not master or layout)
2. Click **Insert Footnote** in the sidebar
3. A footnote text box will be created at the bottom of the slide
4. Replace the placeholder text with your actual footnotes

**What gets created:**
- **1 text box** positioned at the bottom of the slide
- Default text: "Footnotes: 1. Footnotes use 9-point text 2. Items should be numbered and presented without commas."
- Font size: 9pt
- Fully editable after insertion

**Technical details:**

- Position: 5.05 inches from top, 0.33 inches from left
- Width: 9.34 inches (right edge at 9.67", leaving 0.33" from right edge)
- Height: 0.5 inches starting height
- Font size: 9pt (smaller than body text for visual hierarchy)
- Internal padding: 0.01 inches on all sides (top, bottom, left, right)
- Aligns perfectly with column layouts (same 0.33" left/right margins)
- Auto-fit: Not programmatically settable (see note below)
- Standard slide dimensions: 10" wide × 5.625" tall

**Use cases:**
- Adding citations or references for data/statistics
- Including disclaimers or legal text
- Providing source attribution
- Adding explanatory notes or context
- Listing assumptions or methodology notes

**Formatting tips:**
- Use numbered format: "1. source, 2. source, 3. source"
- Keep footnotes concise - they should supplement, not dominate
- Consider using a slightly lighter text color for visual separation
- If footnotes are long, increase box height manually

**To enable auto-resize (if needed):**
Google Apps Script cannot programmatically set "Resize shape to fit text". To enable it manually:
1. Click on the footnote text box
2. Right-click and select **Format options**
3. Under **Size & rotation**, click **Text fitting**
4. Select **Resize shape to fit text**

This will make the text box automatically expand/shrink based on content length.

---

## Size Manipulation

Size Manipulation features allow you to make all selected objects match the anchor's dimensions. Unlike alignment features (which move objects), these features **resize objects** to have consistent dimensions.

All size manipulation features require **at least 2 shapes selected** — one acts as the anchor, and the others are resized to match.

### Align Width

**What it does:**  
Makes all selected objects have the same width as the anchor. Position and height remain unchanged.

**Visual example:**

```
BEFORE:                         AFTER:

[Anchor ]                       [Anchor ]
[Obj1    ]               →      [Obj1   ]
[Obj2]                          [Obj2   ]

All objects now have same width as anchor.
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Align Width** in the sidebar
4. All non-anchor shapes are resized to match anchor's width

**Technical details:**

- Only width changes — height and position stay the same
- Uses `element.setWidth(targetWidth)` from Slides API
- Top-left corner remains fixed during resize

**Use cases:**

- Making all buttons the same width
- Creating uniform column widths
- Standardizing card widths in a layout
- Ensuring consistent text box widths

### Align Height

**What it does:**  
Makes all selected objects have the same height as the anchor. Position and width remain unchanged.

**Visual example:**

```
BEFORE:                         AFTER:

┌────────┐                      ┌────────┐
│ Anchor │                      │ Anchor │
└────────┘                      └────────┘
┌──────┐                        ┌────────┐
│ Obj1 │                 →      │ Obj1   │
└──────┘                        └────────┘
┌─────────┐                     ┌────────┐
│  Obj2   │                     │ Obj2   │
└─────────┘                     └────────┘

All objects now have same height as anchor.
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Align Height** in the sidebar
4. All non-anchor shapes are resized to match anchor's height

**Technical details:**

- Only height changes — width and position stay the same
- Uses `element.setHeight(targetHeight)` from Slides API
- Top-left corner remains fixed during resize

**Use cases:**

- Making all rows the same height
- Creating uniform icon heights
- Standardizing image heights
- Ensuring consistent text box heights

### Align Both

**What it does:**  
Makes all selected objects have the same width AND height as the anchor. Position remains unchanged — objects become exact size duplicates of the anchor.

**Visual example:**

```
BEFORE:                         AFTER:

┌─────────┐                     ┌─────────┐
│ Anchor  │                     │ Anchor  │
└─────────┘                     └─────────┘
┌────┐                          ┌─────────┐
│Obj1│                   →      │  Obj1   │
└────┘                          └─────────┘
┌───────────┐                   ┌─────────┐
│   Obj2    │                   │  Obj2   │
└───────────┘                   └─────────┘

All objects now have same width AND height as anchor.
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Align Both** in the sidebar
4. All non-anchor shapes are resized to match anchor's dimensions

**Technical details:**

- Both width AND height change — position stays the same
- Uses both `setWidth()` and `setHeight()` from Slides API
- Top-left corner remains fixed during resize
- All objects become identical in size to the anchor

**Use cases:**

- Creating perfectly uniform elements (buttons, icons, cards)
- Making all photos the same size
- Standardizing shape dimensions across a slide
- Ensuring visual consistency in layouts

---

## Stretch Objects

Stretch Objects features extend one edge of selected objects to align with the anchor's corresponding edge. Unlike Size Manipulation (which changes size but keeps position), stretching **extends one edge** while keeping the opposite edge fixed.

All stretch operations require **at least 2 shapes selected**.

### Stretch Left

**What it does:**  
Extends objects' left edge to match the anchor's left edge. The right edge stays fixed (objects grow/shrink leftward).

**Visual example:**

```
BEFORE:                         AFTER:

┌────────┐                      ┌────────┐
│ Anchor │                      │ Anchor │
└────────┘                      └────────┘
    ┌──────┐                    ┌──────────┐
    │ Obj1 │             →      │   Obj1   │
    └──────┘                    └──────────┘

Obj1's left edge extended to match anchor's left edge.
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Stretch Left** in the sidebar
4. Objects extend leftward until their left edge matches anchor's left edge

**Technical details:**

- Left edge moves to anchor's left position
- Right edge stays fixed (original position)
- Width changes to accommodate the new left edge
- Formula: `newWidth = currentRight - anchorLeft`

**Use cases:**

- Extending backgrounds to a common left boundary
- Creating left-aligned blocks with consistent left edge
- Stretching elements to fill space on the left

### Stretch Right

**What it does:**  
Extends objects' right edge to match the anchor's right edge. The left edge stays fixed (objects grow/shrink rightward).

**Visual example:**

```
BEFORE:                         AFTER:

┌────────┐                      ┌────────┐
│ Anchor │                      │ Anchor │
└────────┘                      └────────┘
┌──────┐                        ┌──────────┐
│ Obj1 │                 →      │   Obj1   │
└──────┘                        └──────────┘

Obj1's right edge extended to match anchor's right edge.
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Stretch Right** in the sidebar
4. Objects extend rightward until their right edge matches anchor's right edge

**Technical details:**

- Right edge moves to anchor's right position
- Left edge stays fixed (original position)
- Width changes to accommodate the new right edge
- Formula: `newWidth = anchorRight - currentLeft`

**Use cases:**

- Extending backgrounds to a common right boundary
- Creating right-aligned blocks with consistent right edge
- Stretching elements to fill space on the right

### Stretch Top

**What it does:**  
Extends objects' top edge to match the anchor's top edge. The bottom edge stays fixed (objects grow/shrink upward).

**Visual example:**

```
BEFORE:                         AFTER:

┌────────┐                      ┌────────┐
│ Anchor │                      │ Anchor │
└────────┘                      └────────┘
                                ┌────────┐
                                │        │
┌────────┐              →       │  Obj1  │
│  Obj1  │                      │        │
└────────┘                      └────────┘

Obj1's top edge extended to match anchor's top edge.
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Stretch Top** in the sidebar
4. Objects extend upward until their top edge matches anchor's top edge

**Technical details:**

- Top edge moves to anchor's top position
- Bottom edge stays fixed (original position)
- Height changes to accommodate the new top edge
- Formula: `newHeight = currentBottom - anchorTop`

**Use cases:**

- Extending backgrounds to a common top boundary
- Creating top-aligned blocks with consistent top edge
- Stretching elements to fill space upward

### Stretch Bottom

**What it does:**  
Extends objects' bottom edge to match the anchor's bottom edge. The top edge stays fixed (objects grow/shrink downward).

**Visual example:**

```
BEFORE:                         AFTER:

┌────────┐                      ┌────────┐
│  Obj1  │                      │  Obj1  │
└────────┘                      │        │
                        →       │        │
┌────────┐                      │        │
│ Anchor │                      └────────┘
└────────┘                      ┌────────┐
                                │ Anchor │
                                └────────┘

Obj1's bottom edge extended to match anchor's bottom edge.
```

**How to use:**

1. Select 2 or more shapes on your slide
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Stretch Bottom** in the sidebar
4. Objects extend downward until their bottom edge matches anchor's bottom edge

**Technical details:**

- Bottom edge moves to anchor's bottom position
- Top edge stays fixed (original position)
- Height changes to accommodate the new bottom edge
- Formula: `newHeight = anchorBottom - currentTop`

**Use cases:**

- Extending backgrounds to a common bottom boundary
- Creating bottom-aligned blocks with consistent bottom edge
- Stretching elements to fill space downward

---

## Fill Space

Fill Space features stretch selected objects to close gaps between them and the anchor. Unlike Stretch Objects (which extend to match edges), Fill Space only works when there is **actual space/gap** between the objects — it extends one edge until it touches the anchor.

All fill space operations require **at least 2 shapes selected** with gaps between them.

### Fill Left

**What it does:**  
Stretches objects to fill the gap between them and the anchor's right edge. The object's left edge extends until it touches the anchor's right edge. Right edge stays fixed.

**Visual example:**

```
BEFORE:                         AFTER:

┌────────┐         ┌──────┐     ┌────────┐┌──────────┐
│ Anchor │   gap   │ Obj1 │     │ Anchor ││   Obj1   │
└────────┘         └──────┘     └────────┘└──────────┘
            ↑                           ↑
         (gap filled)              (no gap)
```

**How to use:**

1. Select 2 or more shapes with gaps between them
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Fill Left** in the sidebar
4. Objects extend leftward to close the gap with the anchor

**Technical details:**

- Only works if object is to the right of anchor (gap exists)
- Left edge moves to anchor's right edge position
- Right edge stays fixed (original position)
- If objects already overlap, no change occurs
- Formula: `newWidth = currentRight - anchorRight`

**Use cases:**

- Filling horizontal space between two objects
- Creating connectors between elements
- Extending backgrounds to meet boundaries
- Closing gaps in layouts

### Fill Right

**What it does:**  
Stretches objects to fill the gap between them and the anchor's left edge. The object's right edge extends until it touches the anchor's left edge. Left edge stays fixed.

**Visual example:**

```
BEFORE:                         AFTER:

┌──────┐         ┌────────┐     ┌──────────┐┌────────┐
│ Obj1 │   gap   │ Anchor │     │   Obj1   ││ Anchor │
└──────┘         └────────┘     └──────────┘└────────┘
            ↑                           ↑
         (gap filled)              (no gap)
```

**How to use:**

1. Select 2 or more shapes with gaps between them
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Fill Right** in the sidebar
4. Objects extend rightward to close the gap with the anchor

**Technical details:**

- Only works if object is to the left of anchor (gap exists)
- Right edge moves to anchor's left edge position
- Left edge stays fixed (original position)
- If objects already overlap, no change occurs
- Formula: `newWidth = anchorLeft - currentLeft`

**Use cases:**

- Filling horizontal space between two objects
- Creating connectors between elements
- Extending backgrounds to meet boundaries
- Closing gaps in layouts

### Fill Top

**What it does:**  
Stretches objects to fill the gap between them and the anchor's bottom edge. The object's top edge extends until it touches the anchor's bottom edge. Bottom edge stays fixed.

**Visual example:**

```
BEFORE:                         AFTER:

┌────────┐                      ┌────────┐
│ Anchor │                      │ Anchor │
└────────┘                      └────────┘
   gap                          ┌────────┐
┌────────┐              →       │        │
│  Obj1  │                      │  Obj1  │
└────────┘                      └────────┘
    ↑                               ↑
 (gap filled)                  (no gap)
```

**How to use:**

1. Select 2 or more shapes with gaps between them
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Fill Top** in the sidebar
4. Objects extend upward to close the gap with the anchor

**Technical details:**

- Only works if object is below anchor (gap exists)
- Top edge moves to anchor's bottom edge position
- Bottom edge stays fixed (original position)
- If objects already overlap, no change occurs
- Formula: `newHeight = currentBottom - anchorBottom`

**Use cases:**

- Filling vertical space between two objects
- Creating vertical connectors
- Extending backgrounds vertically
- Closing vertical gaps in layouts

### Fill Bottom

**What it does:**  
Stretches objects to fill the gap between them and the anchor's top edge. The object's bottom edge extends until it touches the anchor's top edge. Top edge stays fixed.

**Visual example:**

```
BEFORE:                         AFTER:

┌────────┐                      ┌────────┐
│  Obj1  │                      │  Obj1  │
└────────┘                      │        │
   gap                          └────────┘
┌────────┐              →       ┌────────┐
│ Anchor │                      │ Anchor │
└────────┘                      └────────┘
    ↑                               ↑
 (gap filled)                  (no gap)
```

**How to use:**

1. Select 2 or more shapes with gaps between them
2. *(Optional)* Click **Set Anchor** to designate the reference shape
3. Click **Fill Bottom** in the sidebar
4. Objects extend downward to close the gap with the anchor

**Technical details:**

- Only works if object is above anchor (gap exists)
- Bottom edge moves to anchor's top edge position
- Top edge stays fixed (original position)
- If objects already overlap, no change occurs
- Formula: `newHeight = anchorTop - currentTop`

**Use cases:**

- Filling vertical space between two objects
- Creating vertical connectors
- Extending backgrounds vertically
- Closing vertical gaps in layouts

---

## Magic Resizer

Magic Resizer is an interactive feature that scales selected objects by a user-specified percentage. Unlike other sizing features (which use an anchor), Magic Resizer works on **any number of selected objects** and resizes them proportionally.

**What it does:**  
Opens a dialog where you can enter a percentage to resize selected objects. All objects scale proportionally — both width and height are multiplied by the same factor.

**Visual example:**

```
BEFORE (100%):                  AFTER (150%):

┌──────┐                        ┌─────────┐
│ Box1 │                        │  Box1   │
└──────┘                        └─────────┘
  ●                      →         ●
 Icon                            Icon
                                (1.5x size)
```

**How to use:**

1. Select 1 or more objects on your slide
2. Click **Magic Resizer** in the sidebar
3. Enter a percentage in the dialog:
   - `100%` = no change (original size)
   - `50%` = half the current size
   - `200%` = double the current size
   - `120%` = 20% larger
   - `75%` = 25% smaller
4. Click **Apply** to resize

**Technical details:**

- Opens an HTML modal dialog with percentage input
- Accepts any positive number (1% to 1000%+)
- Multiplies both width and height by `percentage / 100`
- Top-left corner of each object stays in the same position
- Works independently on each selected object
- Does NOT require an anchor (anchor-independent feature)

**Examples:**

| Original Size | Percentage | New Size  |
|---------------|------------|-----------|
| 200 × 100 pt  | 50%        | 100 × 50  |
| 200 × 100 pt  | 200%       | 400 × 200 |
| 200 × 100 pt  | 120%       | 240 × 120 |
| 200 × 100 pt  | 75%        | 150 × 75  |

**Use cases:**

- Quickly scaling elements by a precise factor
- Making all selected objects 20% larger or 50% smaller
- Batch resizing multiple objects proportionally
- Fine-tuning element sizes with precise percentages
- Creating size variations (e.g., make duplicates at 75%, 100%, 125%)

**Dialog features:**

- Clean, modern interface with clear instructions
- Input validation (must be positive number)
- Example percentages shown for reference
- Real-time feedback with success/error messages
- Keyboard shortcuts (Enter to apply)
- Auto-closes after successful resize

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
| `swapPositions()` | — | `string` | Swaps the positions of exactly 2 shapes |
| `distributeHorizontally()` | — | `string` | Distributes shapes evenly across horizontal space |
| `distributeVertically()` | — | `string` | Distributes shapes evenly across vertical space |
| `dockLeft()` | — | `string` | Moves shapes until right edges touch anchor's left edge |
| `dockRight()` | — | `string` | Moves shapes until left edges touch anchor's right edge |
| `dockTop()` | — | `string` | Moves shapes until bottom edges touch anchor's top edge |
| `dockBottom()` | — | `string` | Moves shapes until top edges touch anchor's bottom edge |
| `alignToTable(alignment, padding)` | `string`, `number` | `string` | Aligns shapes within table cells |
| `swapTableRows(rowA, rowB, keepFormatting)` | `number`, `number`, `boolean` | `string` | Swaps two table rows (optionally preserving source formatting) |
| `swapTableColumns(colA, colB, keepFormatting)` | `number`, `number`, `boolean` | `string` | Swaps two table columns (optionally preserving source formatting) |
| `arrangeMatrix(rows, cols, spacing)` | `number`, `number`, `number` | `string` | Arranges shapes in grid layout |
| `insertTwoColumns()` | — | `string` | Inserts 2-column layout with title and content boxes |
| `insertThreeColumns()` | — | `string` | Inserts 3-column layout with title and content boxes |
| `insertFourColumns()` | — | `string` | Inserts 4-column layout with title and content boxes |
| `insertFootnote()` | — | `string` | Inserts footnote text box at bottom of slide |
| `alignWidth()` | — | `string` | Makes selected objects have same width as anchor |
| `alignHeight()` | — | `string` | Makes selected objects have same height as anchor |
| `alignBoth()` | — | `string` | Makes selected objects have same width AND height as anchor |
| `stretchLeft()` | — | `string` | Extends objects' left edge to match anchor's left edge |
| `stretchRight()` | — | `string` | Extends objects' right edge to match anchor's right edge |
| `stretchTop()` | — | `string` | Extends objects' top edge to match anchor's top edge |
| `stretchBottom()` | — | `string` | Extends objects' bottom edge to match anchor's bottom edge |
| `fillLeft()` | — | `string` | Stretches objects to fill gap to anchor's right edge |
| `fillRight()` | — | `string` | Stretches objects to fill gap to anchor's left edge |
| `fillTop()` | — | `string` | Stretches objects to fill gap to anchor's bottom edge |
| `fillBottom()` | — | `string` | Stretches objects to fill gap to anchor's top edge |
| `showMagicResizerDialog()` | — | `string` | Opens dialog for percentage-based resizing |
| `applyMagicResize(percentage)` | `number` | `string` | Scales selected objects by percentage factor |

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
| `runSwapPositions()` | Shows "Swapping...", calls `swapPositions()`, shows result |
| `runDistributeHorizontally()` | Shows "Distributing...", calls `distributeHorizontally()`, shows result |
| `runDistributeVertically()` | Shows "Distributing...", calls `distributeVertically()`, shows result |
| `runDockLeft()` | Shows "Docking left...", calls `dockLeft()`, shows result |
| `runDockRight()` | Shows "Docking right...", calls `dockRight()`, shows result |
| `runDockTop()` | Shows "Docking top...", calls `dockTop()`, shows result |
| `runDockBottom()` | Shows "Docking bottom...", calls `dockBottom()`, shows result |
| `showTableAlignDialog()` | Shows modal dialog for table alignment options |
| `hideTableAlignDialog()` | Hides the table alignment modal dialog |
| `runTableAlign()` | Gets selected alignment and padding, calls `alignToTable()` |
| `showTableSwapDialog(mode)` | Shows row/column swap dialog and configures labels |
| `hideTableSwapDialog()` | Hides the row/column swap dialog |
| `runTableSwap()` | Reads indexes + formatting checkbox, calls `swapTableRows()` or `swapTableColumns()` |
| `runArrangeMatrix()` | Gets rows/cols/spacing, calls `arrangeMatrix()`, shows result |
| `runInsertTwoColumns()` | Shows "Inserting...", calls `insertTwoColumns()`, shows result |
| `runInsertThreeColumns()` | Shows "Inserting...", calls `insertThreeColumns()`, shows result |
| `runInsertFourColumns()` | Shows "Inserting...", calls `insertFourColumns()`, shows result |
| `runInsertFootnote()` | Shows "Inserting...", calls `insertFootnote()`, shows result |
| `runAlignWidth()` | Shows "Aligning width...", calls `alignWidth()`, shows result |
| `runAlignHeight()` | Shows "Aligning height...", calls `alignHeight()`, shows result |
| `runAlignBoth()` | Shows "Aligning size...", calls `alignBoth()`, shows result |
| `runStretchLeft()` | Shows "Stretching left...", calls `stretchLeft()`, shows result |
| `runStretchRight()` | Shows "Stretching right...", calls `stretchRight()`, shows result |
| `runStretchTop()` | Shows "Stretching top...", calls `stretchTop()`, shows result |
| `runStretchBottom()` | Shows "Stretching bottom...", calls `stretchBottom()`, shows result |
| `runFillLeft()` | Shows "Filling left...", calls `fillLeft()`, shows result |
| `runFillRight()` | Shows "Filling right...", calls `fillRight()`, shows result |
| `runFillTop()` | Shows "Filling top...", calls `fillTop()`, shows result |
| `runFillBottom()` | Shows "Filling bottom...", calls `fillBottom()`, shows result |
| `runMagicResizer()` | Shows "Opening...", calls `showMagicResizerDialog()`, shows result |

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

### Table Swap Limitations

- **Border swap not supported**: Google Apps Script table APIs used here do not provide practical border transfer support in this flow.
- **Merged cells not supported for swap**: If either target cell is merged, the operation is blocked with a clear error.
- **Single-table resolution rules**:
  - If exactly one table is selected, that table is used.
  - If no table is selected but exactly one table exists on current slide, that table is used.
  - Otherwise, user must select a single target table.

### Matrix Alignment Limitations

- **Selection order matters**: Shapes are arranged in the order they were selected, not sorted by position.
- **Matrix bounds**: The matrix uses the bounding box of current selection as its reference area.
- **Incomplete rows**: If shapes don't perfectly fit the dimensions, the last row may be incomplete (expected behavior).

### Position Manipulation Limitations

**Swap Positions:**
- Requires exactly 2 shapes - no more, no less
- Only swaps position (left, top) - size, rotation, and styling remain unchanged
- Shapes may overlap after swapping if they have different dimensions

**Distribution:**
- Requires at least 3 shapes (2 anchors + at least 1 to distribute)
- If shapes don't fit in available space, they will overlap (same as PowerPoint behavior)
- Leftmost/rightmost (horizontal) or topmost/bottommost (vertical) shapes stay fixed
- Only gaps are equalized - shapes themselves don't resize

**Docking:**
- Requires at least 2 shapes (anchor + shapes to dock)
- Multiple shapes will stack at the same position (overlap) - dock one at a time to avoid
- Only moves in one direction (horizontal for left/right, vertical for top/bottom)
- Perpendicular position (e.g., vertical position for dock left/right) remains unchanged

### Slide Layout Limitations

**Column Layouts:**
- Requires an active slide (won't work on master slides or layouts)
- All column layouts use fixed measurements (title: 0.6", content: 2.95", gap: 0.05")
- Content boxes end at 5.0" from top to leave room for footnotes (5.05")
- Left and right margins: 0.33 inches (9.34" total horizontal space)
- With 4 columns, text boxes are narrow (~2.18") - may need shorter text
- Text boxes are created with placeholder text that must be manually edited
- Internal padding of 0.01" on all sides provides text spacing
- No undo for inserted text boxes - use Ctrl+Z/Cmd+Z immediately after if needed

**Footnote:**
- Creates a single text box at a fixed position (5.05" from top, 0.33" from left)
- Aligns perfectly with column layouts (same left/right margins of 0.33")
- If you already have content at that position, boxes may overlap
- Width is fixed at 9.34 inches - manually resize if needed
- Height starts at 0.5" - does NOT auto-resize (Google Apps Script limitation)
- Internal padding of 0.01" on all sides for text spacing
- To enable auto-resize: manually set "Resize shape to fit text" in Format Options

**General:**
- All slide layout operations insert new text boxes - they don't modify existing shapes
- Text boxes are independent objects and don't automatically resize with content
- Recommended to use on empty or partially empty slides
- Consider using slide masters or templates for frequently used layouts

### Size Manipulation Limitations

**Align Width / Height / Both:**
- Requires at least 2 shapes (anchor + shapes to resize)
- Position remains unchanged - only dimensions change
- Some element types may not support resizing (caught by try/catch)
- Objects become same size as anchor, but may look disproportionate if aspect ratios differ
- Top-left corner stays fixed during resize

**Key differences:**
- **Align Width**: Only width changes (height unchanged)
- **Align Height**: Only height changes (width unchanged)
- **Align Both**: Both dimensions change (may distort if aspect ratios differ)

### Stretch Objects Limitations

**Stretch Left / Right / Top / Bottom:**
- Requires at least 2 shapes (anchor + shapes to stretch)
- One edge moves, opposite edge stays fixed
- Width or height changes to accommodate the stretch
- If stretch would create negative size (edge beyond opposite edge), operation is skipped
- Some element types may not support stretching (caught by try/catch)

**Key behaviors:**
- **Stretch Left**: Extends leftward (right edge fixed)
- **Stretch Right**: Extends rightward (left edge fixed)
- **Stretch Top**: Extends upward (bottom edge fixed)
- **Stretch Bottom**: Extends downward (top edge fixed)

### Fill Space Limitations

**Fill Left / Right / Top / Bottom:**
- Requires at least 2 shapes with actual gaps between them
- Only works when gap exists - if objects already overlap, no change occurs
- Objects extend to touch the anchor, not align edges with anchor
- One edge moves to touch anchor's opposite edge, other edge stays fixed
- Warning message shown if no gaps found to fill

**Key behaviors:**
- **Fill Left**: Extends left to touch anchor's right edge (only if object is right of anchor)
- **Fill Right**: Extends right to touch anchor's left edge (only if object is left of anchor)
- **Fill Top**: Extends top to touch anchor's bottom edge (only if object is below anchor)
- **Fill Bottom**: Extends bottom to touch anchor's top edge (only if object is above anchor)

### Magic Resizer Limitations

- Opens a modal dialog (blocks interaction with slides until closed)
- Accepts any positive percentage (1% to 1000%+)
- Does NOT use anchor system - works independently on each selected object
- Minimum 1 object required (no maximum)
- Top-left corner stays fixed during resize
- Proportional scaling only - cannot resize width and height independently
- Some element types may not support resizing (caught by try/catch)
- Very small percentages (< 5%) may create nearly invisible objects
- Very large percentages (> 500%) may create objects that extend beyond slide bounds

---

## File Structure

```
superslides/
├── src/
│   ├── appsscript.json    # Manifest (scopes, runtime)
│   ├── Code.gs            # Server-side logic (~3331 lines)
│   └── Sidebar.html       # UI + client-side JS (~1292 lines)
└── FEATURES.md            # This documentation
```

---

*Last updated: February 2026*
