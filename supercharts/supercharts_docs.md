# SuperCharts Documentation

**Version:** 0.1.0 (Pre-release)  
**Last Updated:** February 8, 2026  
**Status:** Planning & Initial Development

---

## Table of Contents

1. [Overview](#overview)
2. [Product Vision](#product-vision)
3. [Technical Architecture](#technical-architecture)
4. [Core Features](#core-features)
5. [User Experience](#user-experience)
6. [Implementation Details](#implementation-details)
7. [Development Roadmap](#development-roadmap)
8. [Design Specifications](#design-specifications)
9. [Testing & Quality](#testing--quality)
10. [Deployment & Distribution](#deployment--distribution)
11. [Business Considerations](#business-considerations)
12. [Resources & References](#resources--references)

---

## Overview

### What is SuperCharts?

SuperCharts is a Google Slides add-on that provides advanced charting capabilities similar to ThinkCell or Mekko Graphics, enabling users to create professional, highly customizable charts for business presentations.

### The Problem We're Solving

Google Slides' native charting is frustrating and limited:
- Poor customization options
- No advanced chart types (waterfall, Mekko, Gantt)
- Difficult to format numbers (e.g., turning 19,786,568 into $19.8M)
- No annotation features (CAGR lines, callouts, brackets)
- Constant redirects to sidebar for basic edits
- Not designed for consulting/business presentations

### The Solution

A native Google Slides add-on that:
- Creates charts using native Slides shapes (fully editable)
- Provides ThinkCell-style advanced chart types
- Offers one-click annotations and formatting
- Enables intuitive data input (paste from Excel/Sheets)
- Supports professional number formatting
- Uses context-aware editing (click chart element → relevant controls appear)

---

## Product Vision

### Target Users

**Primary:**
- Management consultants (McKinsey, BCG, Bain, etc.)
- Investment bankers
- Corporate strategy teams
- Business analysts
- MBA students

**Secondary:**
- Product managers
- Sales teams
- Marketing professionals
- Anyone creating business presentations

### Competitive Positioning

| Feature | Google Slides | ThinkCell | SuperCharts |
|---------|---------------|-----------|-------------|
| **Platform** | Native (limited) | PowerPoint only | Google Slides |
| **Price** | Free | $300+/year | TBD (likely $50-100/year) |
| **Waterfall charts** | No | Yes | Yes |
| **Mekko charts** | No | Yes | Yes (roadmap) |
| **CAGR annotations** | Manual | Yes | Yes |
| **Number formatting** | Basic | Advanced | Advanced |
| **Data paste** | Clunky | Excellent | Excellent |
| **Editable charts** | Sort of | Yes | Yes (native shapes) |

### Success Metrics

**Phase 1 (MVP):**
- 100 beta users
- 500 charts created
- 10+ piece of user feedback

**Phase 2 (Launch):**
- 1,000 active users
- 10,000 charts created
- 4.5+ star rating on Marketplace

**Phase 3 (Growth):**
- 10,000+ active users
- $50K+ ARR
- Enterprise customer pilots

---

## Technical Architecture

### Platform Choice: Google Slides Add-on

**Decision:** Google Slides Add-on (NOT Chrome Extension)

**Rationale:**
- Direct access to Slides API (can create/modify shapes programmatically)
- Native integration with Google Workspace
- Cross-browser compatibility
- Mobile support (Slides mobile apps)
- Industry standard approach (how ThinkCell works for PowerPoint)

**Chrome Extension Considered & Rejected:**
- No direct Slides API access
- Would require DOM manipulation hacks
- Chrome-only (excludes Edge, Safari, Firefox)
- No mobile support
- Cannot programmatically create shapes reliably

### Technology Stack

#### Core Technologies
```
Backend/Logic:
├── Google Apps Script (JavaScript-based)
├── Google Slides API
└── Properties Service (data storage)

Frontend/UI:
├── HTML5
├── CSS3 (Tailwind CSS recommended)
├── Vanilla JavaScript (or lightweight framework)
└── HtmlService (for sidebar)

Development Tools:
├── clasp (local development)
├── Git/GitHub (version control)
├── VS Code (code editor)
└── Google Apps Script Editor (backup/testing)
```

#### Why These Choices?

**Google Apps Script:**
- Required for Slides add-ons
- Runs on Google's servers (no backend needed)
- Built-in authentication (via Google accounts)
- Free hosting
- Automatic scaling

**Native Shapes (vs Image Generation):**
- Charts are fully editable after creation
- Users can move/resize individual elements
- Feels native to Slides
- Better user experience
- How ThinkCell works

### Chart Rendering Approach

**Method:** Draw shapes directly using Slides API

**How it works:**
1. User inputs data
2. JavaScript calculates positions/sizes
3. Apps Script creates Slides shapes (rectangles, lines, text boxes)
4. Each element is positioned programmatically
5. Metadata stored in shape properties

**Example:**
```javascript
// A bar chart is literally:
- 4 Rectangle shapes (the bars)
- 4 TextBox shapes (category labels)
- 4 TextBox shapes (value labels)
- 2 Line shapes (axes)
- 1 TextBox shape (title)

Total: 15 individual Slides shapes
```

**Advantages:**
- Fully editable (user can click individual bars)
- Native to Slides
- No image rendering needed
- Fast performance for typical business charts

**Limitations:**
- Complex for 100+ data points (but rare in business charts)
- Must calculate all positioning manually
- More code to write vs using a chart library

---

## Core Features

### MVP Features (Version 0.1)

#### 1. Chart Types
- **Bar Chart** (vertical, horizontal)
  - Single series
  - Custom colors
  - Value labels on bars

#### 2. Data Input
- **Paste from Excel/Sheets** (primary method)
  - Tab-separated values
  - Auto-parse columns
  - Preview before creating
- **Manual Entry** (secondary)
  - Editable table interface
  - Add/remove rows

#### 3. Basic Customization
- **Colors:** Single color or per-bar
- **Labels:** Show/hide values and categories
- **Number Format:** Raw, K, M, custom
- **Axes:** Show/hide Y-axis, gridlines

#### 4. Simple Annotations
- **CAGR Line:** Trend line with growth rate
- **Average Line:** Horizontal reference line

### Version 0.2 Features

#### 5. Advanced Chart Types
- **Waterfall Chart**
  - Start/end bars
  - Floating bars for changes
  - Connecting lines
  - Subtotals

#### 6. Enhanced Annotations
- **Growth Percentage:** Arrows between bars
- **Callout Boxes:** Highlight specific points
- **Range Brackets:** Group time periods
- **Custom Text:** Annotations anywhere

#### 7. Improved Customization
- **Color Schemes:** Pre-defined palettes
- **Fonts:** Custom font selection
- **Sizing:** Precise width/height controls
- **Positioning:** Snap to grid

### Version 0.3 Features

#### 8. Mekko Chart
- Variable width bars
- Two-dimensional data
- Percentage labels

#### 9. Data Linking
- **Import from Google Sheets**
  - Link to Sheet URL
  - Select range
  - Optional auto-update

#### 10. Templates
- **Save chart configurations**
- **Pre-built templates** (consulting style, corporate, minimal)
- **Share templates** with team

### Future Roadmap

- Gantt charts
- Stacked bar charts
- Combo charts (bar + line)
- Export to PowerPoint
- Collaboration features
- AI-powered chart suggestions

---

## User Experience

### Design Philosophy

**Principles:**
1. **Speed:** Common tasks should take seconds, not minutes
2. **Intuition:** No manual needed for basic charts
3. **Power:** Advanced features available but not overwhelming
4. **Professional:** Output looks consultant-grade
5. **Native:** Feels like it belongs in Slides

### User Flows

#### Flow 1: Quick Bar Chart (Primary Use Case)

**Time Target:** 30 seconds

```
1. User opens SuperCharts sidebar
2. User copies data from Excel/Sheets (Ctrl+C)
3. User pastes into textarea
4. SuperCharts auto-parses data
5. Preview shows parsed data (editable)
6. User clicks "Create Chart"
7. Chart appears on current slide
```

**Key Insight:** Most users already have data in spreadsheets. Make paste-and-create the default path.

#### Flow 2: Manual Data Entry

**Time Target:** 2 minutes for 5 data points

```
1. User opens SuperCharts sidebar
2. User clicks "Manual Entry" tab
3. User types labels and values into table
4. User clicks "+ Add Row" for more data
5. User customizes chart style (optional)
6. User clicks "Create Chart"
7. Chart appears on current slide
```

#### Flow 3: Advanced Customization

**Time Target:** 5 minutes

```
1. User creates basic chart (Flow 1 or 2)
2. User clicks on chart element (e.g., Y-axis)
3. Sidebar updates to show relevant controls
4. User adjusts format (e.g., change to $M)
5. Chart updates in real-time
6. User clicks on bar to change color
7. User adds CAGR annotation
8. Final professional chart ready
```

### Interface Layout

#### Sidebar Structure

```
┌─────────────────────────────────────────┐
│  SuperCharts                      [✕]  │
├─────────────────────────────────────────┤
│  [Paste Data] [Manual] [Import Sheet]  │ ← Tabs
├─────────────────────────────────────────┤
│                                         │
│  📊 Data Input Area                    │
│  ┌─────────────────────────────────┐   │
│  │ Paste from Excel or Sheets:     │   │
│  │ ┌─────────────────────────────┐ │   │
│  │ │                             │ │   │
│  │ │                             │ │   │
│  │ └─────────────────────────────┘ │   │
│  │     [Parse Data]                │   │
│  └─────────────────────────────────┘   │
│                                         │
│  📋 Preview                            │
│  ┌─────────────┬───────────────────┐   │
│  │   Label     │      Value        │   │
│  ├─────────────┼───────────────────┤   │
│  │ Q1          │  120              │   │
│  │ Q2          │  150              │   │
│  └─────────────┴───────────────────┘   │
│     [+ Add Row]  [Clear All]            │
│                                         │
│  🎨 Chart Settings                     │
│  Chart Type: [Bar ▼]                   │
│  Color: [#4285F4 🎨]                   │
│  Number Format: [$0.0M ▼]              │
│                                         │
│  ┌─────────────────────────────────┐   │
│  │    [Create Chart] →             │   │
│  └─────────────────────────────────┘   │
│                                         │
│  ℹ️ Tips: Press Ctrl+V to paste data  │
└─────────────────────────────────────────┘
```

### Context-Aware Editing

**Problem:** Google Charts redirects to sidebar for every edit

**Solution:** Smart sidebar that adapts to selected element

**How it works:**

```
User clicks Y-axis
↓
Sidebar instantly shows:
- Number format dropdown
- Min/Max value controls
- Gridline options
- Axis label text

User clicks a bar
↓
Sidebar instantly shows:
- Color picker
- Data value input
- Bar label options

User clicks legend
↓
Sidebar instantly shows:
- Position controls
- Font size
- Show/hide toggle
```

**Implementation:**
- Monitor `SlidesApp.getActivePresentation().getSelection()`
- Read metadata from selected shape
- Dynamically rebuild sidebar content
- Store element type in shape alt-text or custom properties

### Dark Mode Support

**Requirement:** Full dark mode support from Day 1

**Implementation:**
```css
/* Light mode (default) */
body {
  background: #ffffff;
  color: #333333;
}

/* Dark mode */
body.dark-mode {
  background: #1e1e1e;
  color: #e0e0e0;
}
```

**Icons:** Use `currentColor` in SVG so they adapt automatically

**Storage:** Save preference in `localStorage` or Apps Script `PropertiesService`

**Designer deliverable:** Single color palette with light/dark variants

---

## Implementation Details

### Data Flow

```
1. USER INPUT
   ↓
2. PARSE & VALIDATE
   - Check for errors
   - Validate data types
   - Calculate min/max
   ↓
3. CALCULATE LAYOUT
   - Determine bar positions
   - Calculate label positions
   - Scale values to chart height
   ↓
4. GENERATE SHAPES
   - Create rectangles (bars)
   - Create text boxes (labels)
   - Create lines (axes)
   - Store metadata
   ↓
5. INSERT INTO SLIDE
   - Position shapes
   - Apply styling
   - Group related elements
   ↓
6. CHART CREATED ✓
```

### Key Algorithms

#### 1. Bar Height Calculation

```javascript
function calculateBarHeight(value, maxValue, chartHeight) {
  return (value / maxValue) * chartHeight;
}

// Example:
// value = 150, maxValue = 200, chartHeight = 300
// barHeight = (150 / 200) * 300 = 225 pixels
```

#### 2. Bar Positioning

```javascript
function calculateBarPosition(index, barWidth, barSpacing, chartX) {
  return chartX + (index * (barWidth + barSpacing));
}

// Example:
// For bar #2: index=1, barWidth=50, spacing=10, chartX=100
// barX = 100 + (1 * (50 + 10)) = 160 pixels from left
```

#### 3. Number Formatting

```javascript
function formatNumber(value, format) {
  switch(format) {
    case '$0.0M':
      return '$' + (value / 1000000).toFixed(1) + 'M';
    case '$0.0K':
      return '$' + (value / 1000).toFixed(1) + 'K';
    case '0%':
      return (value * 100).toFixed(0) + '%';
    case '0.0%':
      return (value * 100).toFixed(1) + '%';
    default:
      return value.toLocaleString();
  }
}

// Example:
// formatNumber(19786568, '$0.0M') → "$19.8M"
```

#### 4. Smart Data Parsing

```javascript
function smartParseData(text) {
  // Try tab-separated (from Excel/Sheets)
  if (text.includes('\t')) {
    return parseTabDelimited(text);
  }
  
  // Try comma-separated (CSV)
  if (text.includes(',')) {
    return parseCSV(text);
  }
  
  // Try space-separated
  return parseSpaceDelimited(text);
}
```

### Metadata Storage

**Challenge:** How to remember chart configuration after creation?

**Solution:** Store metadata in shape properties

```javascript
// When creating a chart element
const bar = slide.insertShape(/* ... */);

// Store metadata
bar.setDescription(JSON.stringify({
  chartId: 'chart_123',
  elementType: 'bar',
  dataIndex: 0,
  value: 120,
  format: '$0.0M'
}));

// Later, when user clicks the bar
const metadata = JSON.parse(bar.getDescription());
// Now we know: this is bar #0 of chart_123, value 120
```

**What to store:**
- Chart ID (unique identifier)
- Element type (bar, axis, label, annotation)
- Data index (which data point)
- Original value
- Formatting preferences
- Color choices

### Error Handling

**Validation rules:**

```javascript
function validateChartData(data) {
  const errors = [];
  
  // Minimum data points
  if (data.length < 2) {
    errors.push('Need at least 2 data points');
  }
  
  // Empty labels
  data.forEach((item, i) => {
    if (!item.label || item.label.trim() === '') {
      errors.push(`Row ${i+1}: Label cannot be empty`);
    }
  });
  
  // Non-numeric values
  data.forEach((item, i) => {
    if (isNaN(item.value)) {
      errors.push(`Row ${i+1}: "${item.value}" is not a number`);
    }
  });
  
  // Duplicate labels
  const labels = data.map(d => d.label);
  const duplicates = labels.filter((item, index) => 
    labels.indexOf(item) !== index
  );
  if (duplicates.length > 0) {
    errors.push(`Duplicate labels: ${duplicates.join(', ')}`);
  }
  
  return errors;
}
```

**Error display:**
```html
<div class="error-message">
  ⚠️ Please fix these issues:
  <ul>
    <li>Row 3: "abc" is not a number</li>
    <li>Need at least 2 data points</li>
  </ul>
</div>
```

---

## Development Roadmap

### Phase 1: Foundation (Weeks 1-4)

**Goals:**
- Set up development environment
- Build basic bar chart
- Implement data input

**Deliverables:**
- [ ] GitHub repository created
- [ ] clasp configured for local development
- [ ] Basic sidebar HTML/CSS
- [ ] Data input (paste method)
- [ ] Simple bar chart generation
- [ ] Manual testing checklist

**Success Criteria:**
- Can paste data and create basic bar chart
- Chart appears correctly on slide
- Bars are proportional to values

---

### Phase 2: Core Features (Weeks 5-8)

**Goals:**
- Add customization options
- Implement annotations
- Improve UX

**Deliverables:**
- [ ] Color customization
- [ ] Number formatting options
- [ ] CAGR line annotation
- [ ] Average line annotation
- [ ] Manual data entry tab
- [ ] Preview before creating
- [ ] Error validation

**Success Criteria:**
- Charts look professional
- Annotations work correctly
- User can customize colors and formats

---

### Phase 3: Advanced Charts (Weeks 9-12)

**Goals:**
- Add waterfall chart
- Implement advanced annotations
- Polish UX

**Deliverables:**
- [ ] Waterfall chart type
- [ ] Growth percentage arrows
- [ ] Callout boxes
- [ ] Range brackets
- [ ] Context-aware sidebar
- [ ] Dark mode
- [ ] Chart templates

**Success Criteria:**
- Waterfall charts work correctly
- All annotation types functional
- UI feels polished

---

### Phase 4: Beta Testing (Weeks 13-16)

**Goals:**
- Get user feedback
- Fix bugs
- Improve performance

**Deliverables:**
- [ ] Beta program launched
- [ ] 20+ beta users recruited
- [ ] User feedback collected
- [ ] Major bugs fixed
- [ ] Documentation written
- [ ] Demo video created

**Success Criteria:**
- Positive user feedback
- No critical bugs
- Users create 100+ charts

---

### Phase 5: Launch (Week 17+)

**Goals:**
- Public release
- Marketing push
- User acquisition

**Deliverables:**
- [ ] Google Workspace Marketplace listing
- [ ] Landing page (Astro or similar)
- [ ] User guide documentation
- [ ] Social media announcement
- [ ] Product Hunt launch

**Success Criteria:**
- 100+ installs in first week
- 4.0+ star rating
- Featured on Marketplace (goal)

---

## Design Specifications

### Color Palette

**Primary Colors:**
```
Primary Blue:    #4285F4
Success Green:   #34A853
Warning Yellow:  #FBBC04
Error Red:       #EA4335
```

**Neutral Colors:**
```
Light Mode:
- Background:    #FFFFFF
- Surface:       #F5F5F5
- Border:        #E0E0E0
- Text Primary:  #333333
- Text Secondary:#666666

Dark Mode:
- Background:    #1E1E1E
- Surface:       #2D2D2D
- Border:        #404040
- Text Primary:  #E0E0E0
- Text Secondary:#B0B0B0
```

### Typography

```
Headings:
- Font: Google Sans (or system fallback)
- Size: 16px
- Weight: Bold

Body:
- Font: Roboto (or system fallback)
- Size: 14px
- Weight: Regular

Labels:
- Font: Roboto
- Size: 11-12px
- Weight: Regular

Chart Text:
- Font: Roboto or Arial
- Size: 10-14px (configurable)
- Weight: Regular or Bold
```

### Spacing

```
--spacing-xs:  4px
--spacing-sm:  8px
--spacing-md:  16px
--spacing-lg:  24px
--spacing-xl:  32px
```

### Icons

**Requirements for designer:**
- Format: SVG
- Size: 24x24px viewBox
- Color: Use `currentColor` (adapts to light/dark mode)
- Style: Simple, clean, Google Material style
- Stroke width: Consistent (2px)

**Needed icons:**
- Bar chart
- Waterfall chart
- Mekko chart (future)
- CAGR line
- Average line
- Callout
- Color picker
- Format number
- Import data
- Settings
- Dark mode toggle
- Help/info

**Example SVG structure:**
```svg
<svg width="24" height="24" viewBox="0 0 24 24">
  <path d="M..." fill="currentColor"/>
</svg>
```

### Component Specifications

**Buttons:**
```css
Primary Button:
- Background: #4285F4
- Color: #FFFFFF
- Border-radius: 4px
- Padding: 8px 16px
- Font-size: 14px
- Hover: #3367D6

Secondary Button:
- Background: transparent
- Color: #4285F4
- Border: 1px solid #4285F4
- Border-radius: 4px
- Padding: 8px 16px
```

**Input Fields:**
```css
Text Input:
- Border: 1px solid #E0E0E0
- Border-radius: 4px
- Padding: 8px 12px
- Font-size: 14px
- Focus: border-color #4285F4
```

**Dropdowns:**
```css
Select:
- Border: 1px solid #E0E0E0
- Border-radius: 4px
- Padding: 8px 12px
- Background: #FFFFFF
- Arrow: Custom SVG
```

---

## Testing & Quality

### Testing Strategy

**Level 1: Manual Testing (Now)**

Create test spreadsheet with scenarios:

```
Test Case 1: Basic Bar Chart
- Input: 4 data points
- Expected: 4 bars, correctly sized
- Status: ✓ Pass

Test Case 2: Large Numbers
- Input: Values in millions (19786568)
- Format: $0.0M
- Expected: "$19.8M" displayed
- Status: ⚠️ In progress

Test Case 3: Negative Values
- Input: Mix of positive/negative
- Expected: Bars in correct direction
- Status: ⏳ Not tested

Test Case 4: Edge Cases
- Empty data
- One data point
- Very large dataset (100+ points)
- Special characters in labels
```

**Level 2: Automated Testing (Later)**

```javascript
// Unit tests for calculations
function testBarHeightCalculation() {
  const result = calculateBarHeight(150, 200, 300);
  assertEqual(result, 225, 'Bar height incorrect');
}

function testNumberFormatting() {
  const result = formatNumber(19786568, '$0.0M');
  assertEqual(result, '$19.8M', 'Number format incorrect');
}
```

**Level 3: User Acceptance Testing (Beta)**

- Recruit 20+ beta users
- Provide test scenarios
- Collect feedback via form
- Monitor actual usage patterns

### Quality Checklist

**Before Each Release:**

```
Code Quality:
[ ] No console.log statements in production
[ ] All functions have comments
[ ] Error handling in place
[ ] Code reviewed (even solo - read through)

Functionality:
[ ] All features work as expected
[ ] No broken buttons/links
[ ] Charts render correctly
[ ] Data validation works

UX/UI:
[ ] Responsive layout
[ ] Dark mode works
[ ] Icons display correctly
[ ] Loading states shown
[ ] Error messages clear

Performance:
[ ] Charts create in <2 seconds
[ ] Sidebar loads quickly
[ ] No memory leaks
[ ] Works with 20+ data points

Documentation:
[ ] README updated
[ ] CHANGELOG updated
[ ] Code comments added
[ ] User guide current
```

---

## Deployment & Distribution

### Deployment Process

**Local Development → Testing → Production**

```bash
# 1. Develop locally
clasp push

# 2. Test in Apps Script
# (Manual testing in Google Slides)

# 3. Version the release
git tag v0.1.0
git push --tags

# 4. Deploy to Apps Script
clasp deploy --description "v0.1.0 - Initial beta"

# 5. Update Marketplace listing
# (If already published)
```

### Google Workspace Marketplace

**Listing Requirements:**

```
Basic Info:
- Name: SuperCharts
- Tagline: "Professional charts for Google Slides"
- Description: 500-1000 words
- Category: Productivity
- Icon: 128x128px PNG
- Screenshots: 5-10 images (1280x800px)

Privacy & Security:
- Privacy policy URL
- Terms of service URL
- Support email
- OAuth scopes (what permissions needed)

Pricing:
- Free (initially)
- Premium tier (future): $50-100/year
```

**OAuth Scopes Needed:**

```javascript
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/presentations.currentonly",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

**Explanation:**
- `presentations.currentonly`: Access the current presentation (create shapes)
- `script.container.ui`: Show sidebar UI

### Version Numbering

**Semantic Versioning:** MAJOR.MINOR.PATCH

```
v0.1.0 - MVP (bar charts, basic features)
v0.2.0 - Waterfall charts added
v0.2.1 - Bug fix release
v0.3.0 - Mekko charts added
v1.0.0 - Full public release
```

---

## Business Considerations

### Monetization Strategy

**Phase 1: Free (Year 1)**
- Build user base
- Get feedback
- Establish product-market fit

**Phase 2: Freemium (Year 2)**
```
Free Tier:
- Bar charts
- Basic annotations
- Up to 10 charts/month

Pro Tier ($50-100/year):
- All chart types (waterfall, Mekko, Gantt)
- Unlimited charts
- Advanced annotations
- Priority support
- Templates library
```

**Phase 3: Enterprise (Year 3)**
```
Team Plan ($200-500/year):
- Everything in Pro
- Team templates
- Admin dashboard
- SSO integration
- Dedicated support
```

### Revenue Projections

**Conservative Estimate (Year 2):**
```
1,000 free users
100 Pro users @ $75/year = $7,500 ARR
5 Team plans @ $300/year = $1,500 ARR

Total: $9,000 ARR
```

**Optimistic Estimate (Year 3):**
```
10,000 free users
1,000 Pro users @ $75/year = $75,000 ARR
50 Team plans @ $300/year = $15,000 ARR

Total: $90,000 ARR
```

### Compliance & Security

**Current Status:**
- Not SOC 2 compliant (not needed yet)
- Inherits Google's security (Apps Script runs on Google Cloud)
- No customer data stored outside Google Workspace

**When to Consider SOC 2:**
- Revenue > $500K
- Enterprise customers requesting it
- Competitive requirement

**Estimated Cost:**
- Initial certification: $50K-100K
- Annual renewal: $30K-50K

**Alternative:** Leverage Google's SOC 2 compliance + security questionnaire

### Legal Documents Needed

```
Required:
[ ] Privacy Policy
[ ] Terms of Service
[ ] Data Processing Agreement (for enterprise)

Optional (later):
[ ] SLA (Service Level Agreement)
[ ] DPA addendum
[ ] BAA (if healthcare customers)
```

---

## Resources & References

### Official Documentation

- **Google Apps Script:** https://developers.google.com/apps-script
- **Slides API:** https://developers.google.com/slides
- **Workspace Add-ons:** https://developers.google.com/workspace/add-ons
- **clasp:** https://github.com/google/clasp
- **Marketplace Publishing:** https://developers.google.com/workspace/marketplace

### Inspiration & Competitors

- **ThinkCell:** https://www.think-cell.com (PowerPoint only)
- **Mekko Graphics:** https://www.mekkographics.com
- **Lucidchart:** https://www.lucidchart.com (different approach)
- **Beautiful.ai:** https://www.beautiful.ai (AI-powered)

### Development Resources

**Tutorials:**
- Apps Script Quickstart for Slides
- Building Workspace Add-ons Guide
- HTML Service Best Practices

**Tools:**
- **Tailwind CSS:** https://tailwindcss.com (for sidebar styling)
- **Material Icons:** https://fonts.google.com/icons (placeholder icons)
- **Astro:** https://astro.build (for landing page)

### Community

- **r/GoogleAppsScript** on Reddit
- **Stack Overflow** (tag: google-apps-script)
- **Google Apps Script Community** on Google Groups

---

## Appendix

### File Structure (Initial)

```
supercharts/
├── src/
│   ├── Code.gs              # Main Apps Script entry point
│   ├── Sidebar.html         # Sidebar UI
│   ├── BarChart.gs          # Bar chart generation logic
│   ├── WaterfallChart.gs    # Waterfall chart logic (v0.2)
│   ├── Annotations.gs       # Annotation functions
│   ├── DataParser.gs        # Parse pasted data
│   ├── NumberFormatter.gs   # Format numbers
│   └── Utils.gs             # Utility functions
├── assets/
│   ├── icons/               # SVG icons
│   └── screenshots/         # For Marketplace
├── docs/
│   ├── README.md
│   ├── USER_GUIDE.md
│   └── API.md
├── tests/
│   └── TestCases.md
├── .clasp.json              # clasp configuration
├── appsscript.json          # Apps Script manifest
├── package.json             # npm dependencies (if any)
└── README.md