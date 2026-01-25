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
