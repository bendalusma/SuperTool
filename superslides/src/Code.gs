/**
 * ============================================================================
 * SuperSlides - Code.gs
 * ============================================================================
 * 
 * This is the main server-side code for the SuperSlides add-on.
 * All functions here run on Google's servers, NOT in the browser.
 * 
 * The sidebar (Sidebar.html) calls these functions using google.script.run.
 * 
 * Key concepts:
 * - PageElement: Any object on a slide (shape, image, text box, line, etc.)
 * - Anchor: The reference element that other elements align/size to
 * - DocumentProperties: Key-value storage that persists per presentation
 * 
 * ============================================================================
 */


// ============================================================================
// MENU & SIDEBAR SETUP
// ============================================================================

/**
 * onOpen()
 *
 * This is a special "trigger" function that Google Apps Script recognizes.
 * It runs automatically every time someone opens the Slides presentation.
 * 
 * What it does:
 * 1. Gets the Slides UI (menu bar, dialogs, etc.)
 * 2. Creates a new menu called "SuperSlides"
 * 3. Adds a menu item that calls showSidebar() when clicked
 * 4. Attaches the menu to the UI
 */
function onOpen() {
  SlidesApp.getUi()                          // Get the Slides user interface
    .createMenu('SuperSlides')               // Create a new menu with this name
    .addItem('Open Sidebar', 'showSidebar')  // Add item: label + function to call
    .addToUi();                              // Attach the menu to the UI
}

/**
 * showSidebar()
 *
 * Creates and displays the sidebar panel on the right side of Slides.
 * 
 * How it works:
 * 1. HtmlService.createHtmlOutputFromFile() reads Sidebar.html from your project
 * 2. setTitle() sets the text shown at the top of the sidebar
 * 3. showSidebar() displays it in the Slides UI
 * 
 * Note: The sidebar runs in an iframe, so it's sandboxed from the main page.
 * That's why we use google.script.run to communicate between sidebar and server.
 */
function showSidebar() {
  // Load the HTML file and create an HtmlOutput object
  const html = HtmlService
    .createHtmlOutputFromFile('Sidebar')  // Loads Sidebar.html from project files
    .setTitle('SuperSlides');             // Sets the sidebar header text

  // Display the sidebar in the Slides UI
  SlidesApp.getUi().showSidebar(html);
}


// ============================================================================
// ANCHOR MANAGEMENT
// ============================================================================
// 
// The "anchor" is the reference element that other shapes align to.
// For example, if you want to align 3 shapes to a specific rectangle,
// you'd set that rectangle as the anchor first.
//
// We store the anchor's object ID (a unique string) in DocumentProperties.
// This means the anchor persists even if you close and reopen the sidebar.
// ============================================================================

/**
 * setAnchor()
 *
 * Saves the currently selected element as the anchor.
 * 
 * How it works:
 * 1. Get the current selection from the presentation
 * 2. Check if the selection contains page elements (shapes, images, etc.)
 * 3. Take the FIRST selected element's object ID
 * 4. Store that ID in DocumentProperties for later use
 * 
 * Why the FIRST element?
 * - If user selects multiple shapes, we need a deterministic choice
 * - First element is predictable: user selects one shape, clicks Set Anchor
 * 
 * @returns {string} Status message for the sidebar to display
 */
function setAnchor() {
  // Get the current selection state from the active presentation
  const selection = SlidesApp.getActivePresentation().getSelection();

  // getPageElementRange() returns the selected shapes/images/etc.
  // It returns null if:
  // - Nothing is selected
  // - Something other than page elements is selected (like text within a shape)
  const range = selection.getPageElementRange();
  if (!range) {
    return 'No selection — please select a shape first.';
  }

  // Get the array of selected PageElement objects
  const elements = range.getPageElements();
  if (!elements || elements.length === 0) {
    return 'No elements selected — please select a shape first.';
  }

  // Get the unique object ID of the first selected element
  // Every element in Slides has a unique ID string (like "g123abc456")
  const anchorId = elements[0].getObjectId();

  // Store the anchor ID in DocumentProperties
  // DocumentProperties persist per document (presentation), not per user
  // This means all users editing the same presentation share the same anchor
  PropertiesService.getDocumentProperties().setProperty('anchorId', anchorId);

  return 'Anchor set.';
}

/**
 * clearAnchor()
 *
 * Removes the stored anchor from DocumentProperties.
 * After this, alignment operations will use the fallback behavior
 * (last element in selection becomes the anchor).
 * 
 * @returns {string} Status message for the sidebar to display
 */
function clearAnchor() {
  // deleteProperty() removes the key entirely from storage
  // If the key doesn't exist, this does nothing (no error)
  PropertiesService.getDocumentProperties().deleteProperty('anchorId');
  return 'Anchor cleared.';
}

/**
 * getAnchorStatus()
 *
 * Returns a human-readable string describing the current anchor state.
 * Called by the sidebar when it loads to show initial status.
 * 
 * @returns {string} "Anchor is set." or "No anchor set."
 */
function getAnchorStatus() {
  // getProperty() returns null if the key doesn't exist
  const anchorId = PropertiesService.getDocumentProperties().getProperty('anchorId');
  
  // Ternary operator: condition ? valueIfTrue : valueIfFalse
  return anchorId ? 'Anchor is set.' : 'No anchor set.';
}


// ============================================================================
// ANCHOR RESOLUTION (CORE LOGIC)
// ============================================================================

/**
 * getAnchorOrFallback(elements)
 *
 * This is the CORE function that determines which element acts as the anchor.
 * All alignment and sizing operations call this to find their reference element.
 * 
 * Resolution rules (MVP):
 * 1. If an anchor was explicitly set (setAnchor), AND that anchor is in the
 *    current selection, use it.
 * 2. Otherwise, fall back to the LAST element in the selection array.
 * 
 * Why check if anchor is in selection?
 * - User might have set an anchor on a different slide
 * - The anchor shape might have been deleted
 * - We only want to use it if it's relevant to current operation
 * 
 * Why last element as fallback?
 * - It's intentionally imperfect (see README)
 * - Google doesn't guarantee selection order, so "last" is somewhat arbitrary
 * - This lets us test without always requiring explicit anchor
 * - Future Chrome Extension could track actual click order
 *
 * @param {PageElement[]} elements - Array of currently selected page elements
 * @returns {PageElement} The element to use as the anchor
 */
function getAnchorOrFallback(elements) {
  // Get stored anchor ID from document properties
  const props = PropertiesService.getDocumentProperties();
  const anchorId = props.getProperty('anchorId');

  // If an anchor ID exists, try to find it in the current selection
  if (anchorId) {
    // Array.find() returns the first element that matches the condition
    // or undefined if no match is found
    const match = elements.find(el => el.getObjectId() === anchorId);
    
    if (match) {
      // Found the anchor in the selection — use it
      return match;
    }
    // Anchor exists but isn't in selection — fall through to fallback
  }

  // Fallback: use the last element in the array
  // Note: This is intentionally imperfect for MVP testing
  return elements[elements.length - 1];
}


// ============================================================================
// ALIGNMENT FUNCTIONS
// ============================================================================
//
// Alignment moves elements so their edges or centers match the anchor's.
// 
// Coordinate system in Google Slides:
// - Origin (0,0) is the TOP-LEFT corner of the slide
// - X increases going RIGHT
// - Y increases going DOWN
// - getLeft() returns distance from left edge of slide to left edge of element
// - getTop() returns distance from top edge of slide to top edge of element
//
// ============================================================================

/**
 * alignLeft()
 *
 * Aligns all selected elements to the LEFT edge of the anchor.
 * 
 * Visual example:
 * 
 *   BEFORE:                    AFTER:
 *   
 *   [Anchor]                   [Anchor]
 *         [Shape1]      →      [Shape1]
 *     [Shape2]                 [Shape2]
 *   
 *   All left edges now match the anchor's left edge.
 * 
 * How it works:
 * 1. Get the anchor's left position (distance from slide's left edge)
 * 2. For each OTHER element, set its left position to match
 * 
 * @returns {string} Status message describing what happened
 */
function alignLeft() {
  // Get the active presentation and current selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  // Get the selected page elements (shapes, images, etc.)
  const range = selection.getPageElementRange();

  // Validate: must have page elements selected
  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();

  // Validate: need at least 2 elements to align
  // (one anchor + at least one element to move)
  if (!elements || elements.length < 2) {
    return 'Select at least 2 shapes to align.';
  }

  // Determine which element is the anchor using our resolution rules
  const anchor = getAnchorOrFallback(elements);
  
  // Get the anchor's left edge position (in points from slide's left edge)
  // This is the X coordinate we want all other elements to match
  const anchorLeft = anchor.getLeft();

  // Track success/failure counts for status message
  let moved = 0;
  let failed = 0;

  // Loop through all selected elements
  elements.forEach(el => {
    // Skip the anchor itself — we don't want to move it
    if (el.getObjectId() === anchor.getObjectId()) {
      return; // 'return' in forEach acts like 'continue' in a for loop
    }

    try {
      // setLeft() moves the element so its left edge is at this X position
      el.setLeft(anchorLeft);
      moved += 1;
    } catch (e) {
      // Some element types might not support setLeft() (rare)
      // Catch the error so one bad element doesn't break the whole operation
      failed += 1;
    }
  });

  // Return appropriate status message
  if (failed > 0) {
    return `Aligned ${moved} element(s). ${failed} element(s) could not be moved.`;
  }

  return `Aligned ${moved} element(s) to the anchor's left edge.`;
}

/**
 * alignRight()
 *
 * Aligns all selected elements' RIGHT edges to the anchor's right edge.
 * 
 * Visual example:
 * 
 *   BEFORE:                    AFTER:
 *   
 *       [Anchor]                   [Anchor]
 *   [Shape1]            →          [Shape1]
 *     [Shape2]                     [Shape2]
 *   
 *   All right edges now match the anchor's right edge.
 * 
 * How it works:
 * 1. Calculate anchor's right edge position: left + width
 * 2. For each element, calculate where its LEFT edge needs to be
 *    so that its RIGHT edge matches the anchor's right edge
 *    Formula: newLeft = anchorRight - elementWidth
 * 
 * Why this formula?
 * - We can only set the LEFT position of an element (setLeft)
 * - If we want the right edge at position X, and the element is W wide:
 *   leftEdge + width = rightEdge
 *   leftEdge = rightEdge - width
 * 
 * @returns {string} Status message describing what happened
 */
function alignRight() {
  // Get the active presentation and current selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const range = selection.getPageElementRange();

  // Validate: must have page elements selected
  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();

  // Validate: need at least 2 elements
  if (!elements || elements.length < 2) {
    return 'Select at least 2 shapes to align.';
  }

  // Determine the anchor element
  const anchor = getAnchorOrFallback(elements);
  
  // Calculate the anchor's right edge position
  // Right edge = left edge + width
  const anchorRight = anchor.getLeft() + anchor.getWidth();

  // Track success/failure counts
  let moved = 0;
  let failed = 0;

  // Loop through all selected elements
  elements.forEach(el => {
    // Skip the anchor itself
    if (el.getObjectId() === anchor.getObjectId()) {
      return;
    }

    try {
      // Calculate where this element's left edge needs to be
      // so its right edge aligns with the anchor's right edge
      // 
      // Example:
      // - anchorRight = 500 (anchor's right edge is 500 points from left)
      // - el.getWidth() = 100 (this element is 100 points wide)
      // - newLeft = 500 - 100 = 400
      // - Element's left edge at 400, width 100, so right edge at 500 ✓
      const newLeft = anchorRight - el.getWidth();
      el.setLeft(newLeft);
      moved += 1;
    } catch (e) {
      failed += 1;
    }
  });

  // Return status message
  if (failed > 0) {
    return `Aligned ${moved} element(s). ${failed} element(s) could not be moved.`;
  }

  return `Aligned ${moved} element(s) to the anchor's right edge.`;
}

/**
 * alignTop()
 *
 * Aligns all selected elements to the TOP edge of the anchor.
 * 
 * Visual example:
 * 
 *   BEFORE:                    AFTER:
 *   
 *   [Anchor]    [Shape1]       [Anchor]    [Shape1]    [Shape2]
 *                                 
 *               [Shape2]       (All top edges now aligned horizontally)
 *   
 * How it works:
 * 1. Get the anchor's top position (distance from slide's top edge)
 * 2. For each OTHER element, set its top position to match
 * 
 * Note: This is the vertical equivalent of alignLeft().
 * - alignLeft() aligns left edges (horizontal alignment, shapes stack vertically)
 * - alignTop() aligns top edges (vertical alignment, shapes spread horizontally)
 * 
 * Coordinate reminder:
 * - getTop() returns the Y position (distance from slide's TOP edge)
 * - Y increases going DOWN (0 is at the top of the slide)
 * 
 * @returns {string} Status message describing what happened
 */
function alignTop() {
  // Get the active presentation and current selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  // Get the selected page elements (shapes, images, etc.)
  const range = selection.getPageElementRange();

  // Validate: must have page elements selected
  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();

  // Validate: need at least 2 elements to align
  // (one anchor + at least one element to move)
  if (!elements || elements.length < 2) {
    return 'Select at least 2 shapes to align.';
  }

  // Determine which element is the anchor using our resolution rules
  const anchor = getAnchorOrFallback(elements);
  
  // Get the anchor's top edge position (in points from slide's top edge)
  // This is the Y coordinate we want all other elements to match
  const anchorTop = anchor.getTop();

  // Track success/failure counts for status message
  let moved = 0;
  let failed = 0;

  // Loop through all selected elements
  elements.forEach(el => {
    // Skip the anchor itself — we don't want to move it
    if (el.getObjectId() === anchor.getObjectId()) {
      return; // 'return' in forEach acts like 'continue' in a for loop
    }

    try {
      // setTop() moves the element so its top edge is at this Y position
      // The element moves vertically; its horizontal position (left) stays the same
      el.setTop(anchorTop);
      moved += 1;
    } catch (e) {
      // Some element types might not support setTop() (rare)
      // Catch the error so one bad element doesn't break the whole operation
      failed += 1;
    }
  });

  // Return appropriate status message
  if (failed > 0) {
    return `Aligned ${moved} element(s). ${failed} element(s) could not be moved.`;
  }

  return `Aligned ${moved} element(s) to the anchor's top edge.`;
}

/**
 * alignBottom()
 *
 * Aligns all selected elements' BOTTOM edges to the anchor's bottom edge.
 * 
 * Visual example:
 * 
 *   BEFORE:                    AFTER:
 *   
 *   [Anchor]                   
 *               [Shape1]       [Anchor]    [Shape1]    [Shape2]
 *   
 *               [Shape2]       (All bottom edges now aligned horizontally)
 *   
 * How it works:
 * 1. Calculate anchor's bottom edge position: top + height
 * 2. For each element, calculate where its TOP edge needs to be
 *    so that its BOTTOM edge matches the anchor's bottom edge
 *    Formula: newTop = anchorBottom - elementHeight
 * 
 * Why this formula?
 * - We can only set the TOP position of an element (setTop)
 * - If we want the bottom edge at position Y, and the element is H tall:
 *   topEdge + height = bottomEdge
 *   topEdge = bottomEdge - height
 * 
 * Note: This is the vertical equivalent of alignRight().
 * 
 * @returns {string} Status message describing what happened
 */
function alignBottom() {
  // Get the active presentation and current selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const range = selection.getPageElementRange();

  // Validate: must have page elements selected
  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();

  // Validate: need at least 2 elements
  if (!elements || elements.length < 2) {
    return 'Select at least 2 shapes to align.';
  }

  // Determine the anchor element
  const anchor = getAnchorOrFallback(elements);
  
  // Calculate the anchor's bottom edge position
  // Bottom edge = top edge + height
  const anchorBottom = anchor.getTop() + anchor.getHeight();

  // Track success/failure counts
  let moved = 0;
  let failed = 0;

  // Loop through all selected elements
  elements.forEach(el => {
    // Skip the anchor itself
    if (el.getObjectId() === anchor.getObjectId()) {
      return;
    }

    try {
      // Calculate where this element's top edge needs to be
      // so its bottom edge aligns with the anchor's bottom edge
      // 
      // Example:
      // - anchorBottom = 400 (anchor's bottom edge is 400 points from top)
      // - el.getHeight() = 80 (this element is 80 points tall)
      // - newTop = 400 - 80 = 320
      // - Element's top edge at 320, height 80, so bottom edge at 400 ✓
      const newTop = anchorBottom - el.getHeight();
      el.setTop(newTop);
      moved += 1;
    } catch (e) {
      failed += 1;
    }
  });

  // Return status message
  if (failed > 0) {
    return `Aligned ${moved} element(s). ${failed} element(s) could not be moved.`;
  }

  return `Aligned ${moved} element(s) to the anchor's bottom edge.`;
}

/**
 * alignCenter()
 *
 * Aligns all selected elements' HORIZONTAL CENTERS to the anchor's horizontal center.
 * 
 * Visual example:
 * 
 *   BEFORE:                    AFTER:
 *   
 *       [Anchor]                   [Anchor]
 *   [Shape1]            →          [Shape1]
 *     [Shape2]                     [Shape2]
 *   
 *   All horizontal centers now align vertically (shapes are centered on same X-axis).
 * 
 * How it works:
 * 1. Calculate anchor's horizontal center: left + (width / 2)
 * 2. For each element, calculate where its LEFT edge needs to be
 *    so that its CENTER aligns with the anchor's center
 *    Formula: newLeft = anchorCenterX - (elementWidth / 2)
 * 
 * Why this formula?
 * - We can only set the LEFT position of an element (setLeft)
 * - If we want the center at position X, and the element is W wide:
 *   center = leftEdge + (width / 2)
 *   leftEdge = center - (width / 2)
 * 
 * Note: This aligns CENTERS horizontally, so shapes move left/right.
 * This is different from edge-based alignment (left/right edges).
 * 
 * @returns {string} Status message describing what happened
 */
function alignCenter() {
  // Get the active presentation and current selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const range = selection.getPageElementRange();

  // Validate: must have page elements selected
  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();

  // Validate: need at least 2 elements
  if (!elements || elements.length < 2) {
    return 'Select at least 2 shapes to align.';
  }

  // Determine the anchor element
  const anchor = getAnchorOrFallback(elements);
  
  // Calculate the anchor's horizontal center position
  // Center X = left edge + half of width
  const anchorCenterX = anchor.getLeft() + (anchor.getWidth() / 2);

  // Track success/failure counts
  let moved = 0;
  let failed = 0;

  // Loop through all selected elements
  elements.forEach(el => {
    // Skip the anchor itself
    if (el.getObjectId() === anchor.getObjectId()) {
      return;
    }

    try {
      // Calculate where this element's left edge needs to be
      // so its horizontal center aligns with the anchor's center
      // 
      // Example:
      // - anchorCenterX = 300 (anchor's center is 300 points from left)
      // - el.getWidth() = 100 (this element is 100 points wide)
      // - newLeft = 300 - (100 / 2) = 300 - 50 = 250
      // - Element's left edge at 250, width 100, so center at 300 ✓
      const newLeft = anchorCenterX - (el.getWidth() / 2);
      el.setLeft(newLeft);
      moved += 1;
    } catch (e) {
      failed += 1;
    }
  });

  // Return status message
  if (failed > 0) {
    return `Aligned ${moved} element(s). ${failed} element(s) could not be moved.`;
  }

  return `Aligned ${moved} element(s) to the anchor's horizontal center.`;
}

/**
 * alignMiddle()
 *
 * Aligns all selected elements' VERTICAL CENTERS to the anchor's vertical center.
 * 
 * Visual example:
 * 
 *   BEFORE:                    AFTER:
 *   
 *   [Anchor]                   [Anchor]
 *         [Shape1]      →      [Shape1]
 *     [Shape2]                 [Shape2]
 *   
 *   All vertical centers now align horizontally (shapes are centered on same Y-axis).
 * 
 * How it works:
 * 1. Calculate anchor's vertical center: top + (height / 2)
 * 2. For each element, calculate where its TOP edge needs to be
 *    so that its CENTER aligns with the anchor's center
 *    Formula: newTop = anchorCenterY - (elementHeight / 2)
 * 
 * Why this formula?
 * - We can only set the TOP position of an element (setTop)
 * - If we want the center at position Y, and the element is H tall:
 *   center = topEdge + (height / 2)
 *   topEdge = center - (height / 2)
 * 
 * Note: This aligns CENTERS vertically, so shapes move up/down.
 * This is different from edge-based alignment (top/bottom edges).
 * 
 * @returns {string} Status message describing what happened
 */
function alignMiddle() {
  // Get the active presentation and current selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const range = selection.getPageElementRange();

  // Validate: must have page elements selected
  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();

  // Validate: need at least 2 elements
  if (!elements || elements.length < 2) {
    return 'Select at least 2 shapes to align.';
  }

  // Determine the anchor element
  const anchor = getAnchorOrFallback(elements);
  
  // Calculate the anchor's vertical center position
  // Center Y = top edge + half of height
  const anchorCenterY = anchor.getTop() + (anchor.getHeight() / 2);

  // Track success/failure counts
  let moved = 0;
  let failed = 0;

  // Loop through all selected elements
  elements.forEach(el => {
    // Skip the anchor itself
    if (el.getObjectId() === anchor.getObjectId()) {
      return;
    }

    try {
      // Calculate where this element's top edge needs to be
      // so its vertical center aligns with the anchor's center
      // 
      // Example:
      // - anchorCenterY = 200 (anchor's center is 200 points from top)
      // - el.getHeight() = 60 (this element is 60 points tall)
      // - newTop = 200 - (60 / 2) = 200 - 30 = 170
      // - Element's top edge at 170, height 60, so center at 200 ✓
      const newTop = anchorCenterY - (el.getHeight() / 2);
      el.setTop(newTop);
      moved += 1;
    } catch (e) {
      failed += 1;
    }
  });

  // Return status message
  if (failed > 0) {
    return `Aligned ${moved} element(s). ${failed} element(s) could not be moved.`;
  }

  return `Aligned ${moved} element(s) to the anchor's vertical center.`;
}


// ============================================================================
// TABLE-BASED ALIGNMENT
// ============================================================================
//
// These functions align shapes within their containing table cells.
// 
// How it works:
// 1. Find which table cell contains each shape's center point
// 2. Group shapes by their containing cell
// 3. Align shapes within each cell with configurable padding
//
// ============================================================================

/**
 * findContainingCell(shape, slide)
 *
 * Finds which table cell (if any) contains the shape's center point.
 * Returns both the cell and its bounds for efficient processing.
 * 
 * How it works:
 * 1. Get all tables on the slide
 * 2. For each table, iterate through all cells
 * 3. Calculate each cell's bounds (position and size)
 * 4. Check if the shape's center point falls within the cell bounds
 * 5. Return an object with cell and bounds, or null if none found
 * 
 * Why center point?
 * - More predictable than checking if shape overlaps cell
 * - User can visually see where shape is "centered" in cell
 * - Avoids ambiguity when shape spans multiple cells
 *
 * @param {PageElement} shape - The shape to check
 * @param {Slide} slide - The slide containing the shape
 * @returns {Object|null} Object with {cell, left, top, width, height} or null if not found
 */
function findContainingCell(shape, slide) {
  // Get all tables on the slide
  const tables = slide.getTables();
  
  if (!tables || tables.length === 0) {
    return null;
  }

  // Calculate the shape's center point
  const shapeCenterX = shape.getLeft() + (shape.getWidth() / 2);
  const shapeCenterY = shape.getTop() + (shape.getHeight() / 2);

  // Check each table
  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    
    // Get table bounds
    const tableLeft = table.getLeft();
    const tableTop = table.getTop();
    
    // Get table dimensions (rows and columns)
    const numRows = table.getNumRows();
    const numCols = table.getNumColumns();
    
    // Calculate table width by summing column widths
    // (Tables don't have .getWidth() method, so we need to sum column widths)
    let tableWidth = 0;
    for (let col = 0; col < numCols; col++) {
      tableWidth += table.getColumn(col).getWidth();
    }
    
    // Calculate table height by summing row heights
    let tableHeight = 0;
    for (let row = 0; row < numRows; row++) {
      tableHeight += table.getRow(row).getMinimumHeight();
    }
    
    // Calculate cell dimensions
    const cellWidth = tableWidth / numCols;
    const cellHeight = tableHeight / numRows;
    
    // Check each cell in the table
    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        // Calculate cell bounds
        const cellLeft = tableLeft + (col * cellWidth);
        const cellTop = tableTop + (row * cellHeight);
        const cellRight = cellLeft + cellWidth;
        const cellBottom = cellTop + cellHeight;
        
        // Check if shape center is within cell bounds
        if (shapeCenterX >= cellLeft && 
            shapeCenterX < cellRight &&
            shapeCenterY >= cellTop && 
            shapeCenterY < cellBottom) {
          // Found the containing cell! Return cell and bounds
          return {
            cell: table.getCell(row, col),
            left: cellLeft,
            top: cellTop,
            width: cellWidth,
            height: cellHeight
          };
        }
      }
    }
  }
  
  // Shape not contained in any cell
  return null;
}

/**
 * alignShapesInCell(shapes, cellLeft, cellTop, cellWidth, cellHeight, alignment, padding)
 *
 * Aligns multiple shapes within a single table cell.
 * 
 * If multiple shapes are in the same cell, they are stacked vertically
 * (top to bottom) with the specified alignment.
 * 
 * @param {PageElement[]} shapes - Array of shapes to align
 * @param {number} cellLeft - Left edge of the cell
 * @param {number} cellTop - Top edge of the cell
 * @param {number} cellWidth - Width of the cell
 * @param {number} cellHeight - Height of the cell
 * @param {string} alignment - 'left', 'right', or 'center'
 * @param {number} padding - Padding from cell edge (in points)
 * @returns {number} Number of shapes successfully aligned
 */
function alignShapesInCell(shapes, cellLeft, cellTop, cellWidth, cellHeight, alignment, padding) {
  if (!shapes || shapes.length === 0) {
    return 0;
  }

  // Calculate total height of all shapes (including gaps)
  let totalHeight = 0;
  const shapeGap = 5; // Gap between shapes when stacked
  shapes.forEach((shape, index) => {
    totalHeight += shape.getHeight();
    if (index < shapes.length - 1) {
      totalHeight += shapeGap; // Add gap between shapes (not after last one)
    }
  });
  
  // Start position: center vertically within the cell
  let currentTop = cellTop + (cellHeight / 2) - (totalHeight / 2);
  
  let aligned = 0;

  shapes.forEach(shape => {
    try {
      // Calculate horizontal position based on alignment
      let newLeft;
      
      if (alignment === 'left') {
        // Align to left edge with padding
        newLeft = cellLeft + padding;
      } else if (alignment === 'right') {
        // Align to right edge with padding
        newLeft = cellLeft + cellWidth - shape.getWidth() - padding;
      } else { // 'center'
        // Center horizontally
        newLeft = cellLeft + (cellWidth / 2) - (shape.getWidth() / 2);
      }
      
      // Set position
      shape.setLeft(newLeft);
      shape.setTop(currentTop);
      
      // Move down for next shape (add shape height + gap)
      currentTop += shape.getHeight() + shapeGap;
      aligned += 1;
    } catch (e) {
      // Skip shapes that can't be moved
    }
  });

  return aligned;
}

/**
 * alignToTable(alignment, padding)
 *
 * Aligns selected shapes within their containing table cells.
 * 
 * How it works:
 * 1. Get all selected shapes
 * 2. Get the current slide
 * 3. For each shape, find which table cell contains it (by center point)
 * 4. Group shapes by their containing cell
 * 5. For each cell group, align shapes with specified padding
 * 
 * Alignment options:
 * - 'left': Align shapes to left edge of cell with padding
 * - 'right': Align shapes to right edge of cell with padding
 * - 'center': Center shapes horizontally within cell
 * 
 * Multiple shapes in same cell:
 * - Shapes are stacked vertically (top to bottom)
 * - Each shape maintains its alignment (left/right/center)
 * - Small gap (5pt) between stacked shapes
 *
 * @param {string} alignment - 'left', 'right', or 'center'
 * @param {number} padding - Padding from cell edge (in points, default: 10)
 * @returns {string} Status message describing what happened
 */
function alignToTable(alignment, padding) {
  // Validate alignment parameter
  if (!alignment || !['left', 'right', 'center'].includes(alignment)) {
    return 'Invalid alignment. Use "left", "right", or "center".';
  }
  
  // Validate and set default padding
  if (padding === undefined || padding === null) {
    padding = 10; // Default 10 points
  }
  if (padding < 0) {
    padding = 0; // Don't allow negative padding
  }

  // Get the active presentation and current selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const range = selection.getPageElementRange();

  // Validate: must have page elements selected
  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();
  if (!elements || elements.length === 0) {
    return 'No shapes selected.';
  }
  
  // Filter out tables - we only want shapes, not tables themselves
  const shapes = elements.filter(function(element) {
    try {
      const type = element.getPageElementType();
      // Only process SHAPE, IMAGE, TEXT_BOX, etc. - NOT TABLE
      return type !== SlidesApp.PageElementType.TABLE;
    } catch (e) {
      return false; // Skip elements that cause errors
    }
  });
  
  if (!shapes || shapes.length === 0) {
    return 'No shapes selected. Please select shapes inside the table, not the table itself.';
  }

  // Get the current slide directly from selection
  const currentPage = selection.getCurrentPage();
  
  if (!currentPage) {
    return 'Could not determine current page. Please try again.';
  }
  
  // Check if we're on a slide (not a master or layout)
  if (currentPage.getPageType() !== SlidesApp.PageType.SLIDE) {
    return 'Please select shapes on a slide (not master slide or layout).';
  }
  
  const slide = currentPage.asSlide();

  // Check if slide has any tables
  const tables = slide.getTables();
  if (!tables || tables.length === 0) {
    return 'No tables found on this slide. Add a table first.';
  }

  // Group shapes by their containing cell
  // Key: cell identifier (row-col), Value: array of {shape, bounds}
  const cellGroups = {};
  let shapesInCells = 0;
  let shapesNotInCells = 0;

  shapes.forEach(shape => {
    try {
      const cellInfo = findContainingCell(shape, slide);
      
      if (cellInfo) {
        // Create a unique key for this cell (we'll use bounds as identifier)
        const cellKey = `${cellInfo.left.toFixed(2)}-${cellInfo.top.toFixed(2)}`;
        
        if (!cellGroups[cellKey]) {
          cellGroups[cellKey] = {
            bounds: {
              left: cellInfo.left,
              top: cellInfo.top,
              width: cellInfo.width,
              height: cellInfo.height
            },
            shapes: []
          };
        }
        
        cellGroups[cellKey].shapes.push(shape);
        shapesInCells += 1;
      } else {
        shapesNotInCells += 1;
      }
    } catch (e) {
      // Skip shapes that cause errors
      shapesNotInCells += 1;
    }
  });

  // If no shapes are in cells, return early with debug info
  if (shapesInCells === 0) {
    // Add debug information
    const firstShape = shapes[0];
    const shapeInfo = `Shape center: (${(firstShape.getLeft() + firstShape.getWidth()/2).toFixed(1)}, ${(firstShape.getTop() + firstShape.getHeight()/2).toFixed(1)})`;
    const tableInfo = `Tables found: ${tables.length}`;
    
    // Add table bounds info for debugging (with null checks)
    let tableBoundsInfo = '';
    if (tables.length > 0) {
      const table = tables[0];
      const tLeft = table.getLeft();
      const tTop = table.getTop();
      const tRows = table.getNumRows();
      const tCols = table.getNumColumns();
      
      // Calculate table width by summing column widths
      let tWidth = 0;
      for (let col = 0; col < tCols; col++) {
        tWidth += table.getColumn(col).getWidth();
      }
      
      // Calculate table height by summing row heights
      let tHeight = 0;
      for (let row = 0; row < tRows; row++) {
        tHeight += table.getRow(row).getMinimumHeight();
      }
      
      // Check for null values
      if (tLeft === null || tTop === null || tWidth === null || tHeight === null) {
        tableBoundsInfo = `Table: ERROR - Table has null bounds (left=${tLeft}, top=${tTop}, width=${tWidth}, height=${tHeight})`;
      } else {
        tableBoundsInfo = `Table: left=${tLeft.toFixed(1)}, top=${tTop.toFixed(1)}, size=${tWidth.toFixed(1)}x${tHeight.toFixed(1)}, grid=${tRows}x${tCols}`;
        
        // Calculate first cell bounds for reference
        const cellW = tWidth / tCols;
        const cellH = tHeight / tRows;
        tableBoundsInfo += `, Cell[0,0]: (${tLeft.toFixed(1)},${tTop.toFixed(1)}) to (${(tLeft+cellW).toFixed(1)},${(tTop+cellH).toFixed(1)})`;
      }
    }
    
    return `No selected shapes are contained in table cells. ${shapesNotInCells} shape(s) skipped. Debug: ${shapeInfo}, ${tableInfo}, ${tableBoundsInfo}`;
  }

  // Align shapes in each cell group
  let totalAligned = 0;
  let totalFailed = 0;

  Object.keys(cellGroups).forEach(cellKey => {
    const group = cellGroups[cellKey];
    const aligned = alignShapesInCell(
      group.shapes,
      group.bounds.left,
      group.bounds.top,
      group.bounds.width,
      group.bounds.height,
      alignment,
      padding
    );
    
    totalAligned += aligned;
    totalFailed += (group.shapes.length - aligned);
  });

  // Build status message
  let message = `Aligned ${totalAligned} shape(s) within table cells (${alignment} alignment).`;
  
  if (shapesNotInCells > 0) {
    message += ` ${shapesNotInCells} shape(s) not in cells were skipped.`;
  }
  
  if (totalFailed > 0) {
    message += ` ${totalFailed} shape(s) could not be moved.`;
  }

  return message;
}


// ============================================================================
// MATRIX ALIGNMENT
// ============================================================================
//
// These functions arrange shapes into a grid/matrix layout.
// 
// How it works:
// 1. User specifies desired rows and columns
// 2. If shapes don't fit, auto-expand dimensions (e.g., 10 shapes in 3x3 → 4x3)
// 3. Calculate matrix bounds from selection bounds
// 4. Arrange shapes in grid maintaining selection order
//
// ============================================================================

/**
 * calculateMatrixDimensions(numShapes, requestedRows, requestedCols)
 *
 * Calculates the actual matrix dimensions needed to fit all shapes.
 * 
 * If the number of shapes exceeds the requested dimensions, this function
 * auto-expands the rows to accommodate all shapes.
 * 
 * Examples:
 * - 9 shapes, 3x3 requested → returns {rows: 3, cols: 3} (perfect fit)
 * - 10 shapes, 3x3 requested → returns {rows: 4, cols: 3} (auto-expand rows)
 * - 6 shapes, 3x3 requested → returns {rows: 2, cols: 3} (fewer rows needed)
 * 
 * Formula:
 * - actualRows = Math.ceil(numShapes / requestedCols)
 * - actualCols = requestedCols (columns stay fixed)
 * 
 * Why auto-expand rows instead of columns?
 * - More intuitive: shapes flow left-to-right, top-to-bottom
 * - Last row may be incomplete (e.g., 10 shapes in 3 cols = 4 rows, last row has 1)
 *
 * @param {number} numShapes - Total number of shapes to arrange
 * @param {number} requestedRows - User's requested number of rows
 * @param {number} requestedCols - User's requested number of columns
 * @returns {Object} Object with {rows, cols} representing actual dimensions
 */
function calculateMatrixDimensions(numShapes, requestedRows, requestedCols) {
  // Validate inputs
  if (numShapes <= 0) {
    return {rows: 1, cols: 1};
  }
  
  if (requestedRows <= 0 || requestedCols <= 0) {
    return {rows: 1, cols: numShapes};
  }

  // Calculate actual rows needed
  // If shapes fit in requested dimensions, use those
  // Otherwise, expand rows to fit all shapes
  const maxShapesInRequested = requestedRows * requestedCols;
  
  if (numShapes <= maxShapesInRequested) {
    // Shapes fit in requested dimensions
    // But we might need fewer rows if shapes < requested
    const actualRows = Math.ceil(numShapes / requestedCols);
    return {
      rows: actualRows,
      cols: requestedCols
    };
  } else {
    // Auto-expand: calculate rows needed to fit all shapes
    const actualRows = Math.ceil(numShapes / requestedCols);
    return {
      rows: actualRows,
      cols: requestedCols
    };
  }
}

/**
 * getSelectionBounds(elements)
 *
 * Calculates the bounding box that contains all selected shapes.
 * 
 * This bounding box is used as the reference area for the matrix layout.
 * The matrix will be positioned and sized to fit within this area.
 * 
 * How it works:
 * 1. Find the leftmost edge (minimum left)
 * 2. Find the topmost edge (minimum top)
 * 3. Find the rightmost edge (maximum left + width)
 * 4. Find the bottommost edge (maximum top + height)
 * 5. Return bounds object
 *
 * @param {PageElement[]} elements - Array of selected page elements
 * @returns {Object} Object with {left, top, width, height} properties
 */
function getSelectionBounds(elements) {
  if (!elements || elements.length === 0) {
    return {left: 0, top: 0, width: 100, height: 100}; // Default bounds
  }

  // Initialize with first element's bounds
  let minLeft = elements[0].getLeft();
  let minTop = elements[0].getTop();
  let maxRight = elements[0].getLeft() + elements[0].getWidth();
  let maxBottom = elements[0].getTop() + elements[0].getHeight();

  // Find extremes across all elements
  elements.forEach(el => {
    const left = el.getLeft();
    const top = el.getTop();
    const right = left + el.getWidth();
    const bottom = top + el.getHeight();

    if (left < minLeft) minLeft = left;
    if (top < minTop) minTop = top;
    if (right > maxRight) maxRight = right;
    if (bottom > maxBottom) maxBottom = bottom;
  });

  return {
    left: minLeft,
    top: minTop,
    width: maxRight - minLeft,
    height: maxBottom - minTop
  };
}

/**
 * arrangeInMatrix(elements, rows, cols, spacing, bounds)
 *
 * Arranges shapes in a grid/matrix layout within the specified bounds.
 * 
 * How it works:
 * 1. Calculate cell size based on bounds and spacing
 * 2. For each shape (in selection order), calculate its grid position
 * 3. Position shape at calculated cell position
 * 
 * Grid calculation:
 * - cellWidth = (bounds.width - (cols-1)*spacing) / cols
 * - cellHeight = (bounds.height - (rows-1)*spacing) / rows
 * 
 * Position calculation for shape at index i:
 * - row = Math.floor(i / cols)
 * - col = i % cols
 * - left = bounds.left + col * (cellWidth + spacing)
 * - top = bounds.top + row * (cellHeight + spacing)
 * 
 * Note: Shapes maintain their selection order. They are NOT sorted by position.
 *
 * @param {PageElement[]} elements - Array of shapes to arrange
 * @param {number} rows - Number of rows in matrix
 * @param {number} cols - Number of columns in matrix
 * @param {number} spacing - Spacing between shapes (in points)
 * @param {Object} bounds - Bounding box {left, top, width, height}
 * @returns {number} Number of shapes successfully arranged
 */
function arrangeInMatrix(elements, rows, cols, spacing, bounds) {
  if (!elements || elements.length === 0) {
    return 0;
  }

  // Calculate cell dimensions
  // Total spacing = (cols-1) * spacing (spacing between columns)
  const totalHorizontalSpacing = (cols - 1) * spacing;
  const cellWidth = (bounds.width - totalHorizontalSpacing) / cols;
  
  // Total spacing = (rows-1) * spacing (spacing between rows)
  const totalVerticalSpacing = (rows - 1) * spacing;
  const cellHeight = (bounds.height - totalVerticalSpacing) / rows;

  let arranged = 0;

  // Arrange each shape in grid position
  elements.forEach((shape, index) => {
    try {
      // Calculate grid position (row and column)
      const row = Math.floor(index / cols);
      const col = index % cols;
      
      // Calculate actual position
      // For horizontal: start at bounds.left, add column offset
      const left = bounds.left + (col * (cellWidth + spacing));
      
      // For vertical: start at bounds.top, add row offset
      const top = bounds.top + (row * (cellHeight + spacing));
      
      // Center shape within its cell (optional - could align top-left instead)
      // For now, we'll position at top-left of cell
      // User can adjust if needed
      
      shape.setLeft(left);
      shape.setTop(top);
      arranged += 1;
    } catch (e) {
      // Skip shapes that can't be moved
    }
  });

  return arranged;
}

/**
 * arrangeMatrix(rows, cols, spacing)
 *
 * Arranges selected shapes into a grid/matrix layout.
 * 
 * How it works:
 * 1. Get selected shapes
 * 2. Validate inputs (rows, cols, spacing)
 * 3. Calculate actual matrix dimensions (with auto-expand if needed)
 * 4. Get bounding box of current selection
 * 5. Arrange shapes in matrix maintaining selection order
 * 
 * Auto-expansion:
 * - If number of shapes > requested rows × cols, rows are auto-expanded
 * - Example: 10 shapes, 3×3 requested → becomes 4×3 (4 rows, 3 cols)
 * - Last row may be incomplete (e.g., 10 shapes in 3 cols = 4 rows, last row has 1 shape)
 * 
 * Matrix positioning:
 * - Matrix is positioned using the bounding box of current selection
 * - Shapes are arranged within this bounding box area
 * - Spacing is applied between shapes (not around edges)
 *
 * @param {number} rows - Requested number of rows
 * @param {number} cols - Requested number of columns
 * @param {number} spacing - Spacing between shapes (in points, default: 20)
 * @returns {string} Status message describing what happened
 */
function arrangeMatrix(rows, cols, spacing) {
  // Validate inputs
  if (!rows || rows <= 0) {
    return 'Invalid rows. Please specify a positive number.';
  }
  
  if (!cols || cols <= 0) {
    return 'Invalid columns. Please specify a positive number.';
  }
  
  // Validate and set default spacing
  if (spacing === undefined || spacing === null) {
    spacing = 20; // Default 20 points
  }
  if (spacing < 0) {
    spacing = 0; // Don't allow negative spacing
  }

  // Get the active presentation and current selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const range = selection.getPageElementRange();

  // Validate: must have page elements selected
  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();
  if (!elements || elements.length === 0) {
    return 'No shapes selected.';
  }

  const numShapes = elements.length;

  // Calculate actual matrix dimensions (with auto-expand if needed)
  const dimensions = calculateMatrixDimensions(numShapes, rows, cols);
  const actualRows = dimensions.rows;
  const actualCols = dimensions.cols;

  // Build dimension info message
  let dimensionMessage = '';
  if (actualRows !== rows || actualCols !== cols) {
    dimensionMessage = ` (auto-expanded to ${actualRows}×${actualCols} to fit ${numShapes} shape(s))`;
  }

  // Get bounding box of current selection
  const bounds = getSelectionBounds(elements);

  // Arrange shapes in matrix
  const arranged = arrangeInMatrix(elements, actualRows, actualCols, spacing, bounds);

  // Build status message
  let message = `Arranged ${arranged} shape(s) in ${actualRows}×${actualCols} matrix${dimensionMessage}.`;
  
  if (arranged < numShapes) {
    const failed = numShapes - arranged;
    message += ` ${failed} shape(s) could not be moved.`;
  }

  return message;
}
