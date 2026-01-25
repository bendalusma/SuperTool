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
2. [Anchor Management](#anchor-management)
   - [Set Anchor](#set-anchor)
   - [Clear Anchor](#clear-anchor)
   - [Anchor Status](#anchor-status)
3. [How Anchor Resolution Works](#how-anchor-resolution-works)
4. [API Reference](#api-reference)
5. [Known Limitations](#known-limitations)

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

Certain element types (like embedded videos or linked objects) may not support `setLeft()`. These are caught by try/catch and reported in the status message.

---

## File Structure

```
superslides/
├── src/
│   ├── appsscript.json    # Manifest (scopes, runtime)
│   ├── Code.gs            # Server-side logic (~680 lines)
│   └── Sidebar.html       # UI + client-side JS (~380 lines)
└── FEATURES.md            # This documentation
```

---

*Last updated: January 2026*
