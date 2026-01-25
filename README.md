# SuperTools

SuperTools is a productivity suite for Google Workspace focused on fixing gaps in
Google Slides and Google Sheets by adding **precise layout tools, formatting shortcuts,
and advanced charting**.

The project is structured as two sub-brands:

- **SuperSlides** – layout, alignment, sizing, and positioning tools for Google Slides
- **SuperSheets** – formatting, table, and custom formulas 
- **SuperCharts** – advanced charting tools for Google Sheets and Google Slides


The long-term goal is to provide a ThinkCell-like experience for Google Workspace,
starting with a lightweight MVP built using Google Apps Script, with optional future
enhancements via a Chrome Extension.

---

## Project Structure

Each sub-project is developed and deployed independently using `clasp`.

---

## Core Design Principles

- **Fast & predictable**: minimize Google API calls, batch operations when possible
- **Explicit over implicit**: users can explicitly set an anchor shape
- **Progressive enhancement**: MVP works with Apps Script only; advanced UX later
- **Design-first UI**: clean, minimal sidebar with custom icons and styling
- **Modular codebase**: alignment, sizing, positioning logic kept separate

---

## MVP Anchor Logic (Critical Design Decision)

Many features depend on a reference object (“anchor”).

### Anchor Resolution Rules (MVP):
1. If the user explicitly sets an anchor → **always use it**
2. If no anchor is set → fallback to:
3. This fallback is intentionally kept to:
- Evaluate reliability
- Inform future Chrome Extension work

This rule applies to **all alignment, sizing, and positioning features**.

---

## SuperSlides – Planned Features

### 1. Anchor Management
- Set anchor from current selection
- Clear anchor
- Display anchor status in sidebar

---

### 2. Alignment (relative to anchor)
- Align left
- Align right
- Align center (horizontal)
- Align middle (vertical)

---

### 3. Sizing
- Match width to anchor
- Match height to anchor
- Match both width and height

---

### 4. Positioning
- Swap position between two selected objects

---

### 5. Distribution (future)
- Distribute horizontally
- Distribute vertically

---

### 6. Selection Improvements (future)
- Custom selection logic to avoid selecting background or container objects
- More predictable object targeting

---

## SuperSheets – Planned Features

### 1. Number Formatting
- Convert values to:
- `$K`
- `$M`
- `$B`
- Apply formatting to selected cells or ranges

---

### 2. Table Shortcuts
- Add row above
- Add row below
- Copy formatting

---

### 3. Charting (Longer Term)
- Waterfall charts
- Marimekko (Mekko) charts
- Custom formatting options beyond native Sheets charts

---

## UI & Design

- Sidebar-based UI (Google Apps Script limitation)
- Custom buttons with SVG icons (designed in Illustrator / Photoshop)
- Icons embedded via SVG sprite (`icons.html`)
- CSS-based design system (colors, spacing, button states)
- No heavy frontend frameworks (performance-first)

---

## Performance Strategy

- Cache reads (selection, anchor metrics)
- Compute locally, mutate once
- Batch updates using Advanced Slides API when needed
- Avoid per-element API calls inside loops
- Keep sidebar lightweight (no React/Vue)

---

## Future Enhancements (Post-MVP)

### Chrome Extension (Optional Power Mode)
- True keyboard shortcuts (e.g. Ctrl+Alt+L)
- Detect actual last-clicked object
- Track selection order reliably
- Floating UI / toolbar inside Slides canvas

The extension would complement — not replace — the Apps Script add-on.

---

## Development Stack

- Google Apps Script (Slides & Sheets)
- HTML / CSS / Vanilla JavaScript (sidebar)
- `clasp` for local development
- Git + GitHub for version control
- Cursor / VS Code for editing

---

## Status

🚧 **In active development**

Current focus:
- SuperSlides MVP
- Anchor management
- Core alignment & sizing tools