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
// POSITION MANIPULATION
// ============================================================================
//
// These functions manipulate the positions of shapes in various ways:
// - Swap: Exchange positions of two shapes
// - Distribute: Arrange shapes with equal spacing
// - Dock: Move shapes until they touch the anchor's edges
//
// ============================================================================

/**
 * swapPositions()
 *
 * Swaps the positions of exactly two selected shapes.
 * 
 * Visual example:
 * 
 *   BEFORE:                    AFTER:
 *   
 *   [Shape A]                  [Shape B]
 *              [Shape B]  →                 [Shape A]
 *   
 *   Shapes exchange positions while maintaining size and other properties.
 * 
 * How it works:
 * 1. Requires exactly 2 selected elements
 * 2. Stores the left/top coordinates of both shapes
 * 3. Swaps their positions using setLeft() and setTop()
 * 
 * Note: This swaps only position (left, top), not size, rotation, or styling.
 * Shapes may overlap after swapping - this is expected behavior.
 * 
 * Use cases:
 * - Quick position exchange without manual dragging
 * - Swapping placeholder positions in layouts
 * - A/B testing different arrangements
 * 
 * @returns {string} Status message
 */
function swapPositions() {
  // Get the current selection from the presentation
  const selection = SlidesApp.getActivePresentation().getSelection();
  const pageElements = selection.getPageElementRange();

  // Validation: Must have exactly 2 elements
  if (!pageElements) {
    return '⚠️ No elements selected. Please select exactly 2 shapes.';
  }

  const elements = pageElements.getPageElements();
  
  if (elements.length !== 2) {
    return `⚠️ Please select exactly 2 shapes to swap. Currently selected: ${elements.length}`;
  }

  // Get references to both shapes
  const shape1 = elements[0];
  const shape2 = elements[1];

  // Store original positions
  // We need to capture both before making any changes
  const shape1Left = shape1.getLeft();
  const shape1Top = shape1.getTop();
  const shape2Left = shape2.getLeft();
  const shape2Top = shape2.getTop();

  // Perform the swap
  // Shape 1 moves to Shape 2's old position
  shape1.setLeft(shape2Left);
  shape1.setTop(shape2Top);
  
  // Shape 2 moves to Shape 1's old position
  shape2.setLeft(shape1Left);
  shape2.setTop(shape1Top);

  return '✅ Positions swapped successfully!';
}

/**
 * distributeHorizontally()
 *
 * Distributes selected shapes evenly across horizontal space.
 * Similar to PowerPoint's "Distribute Horizontally" feature.
 * 
 * Visual example:
 * 
 *   BEFORE:                         AFTER:
 *   
 *   [1]  [2][3]    [4]       [5]   [1]  [2]  [3]  [4]  [5]
 *                              →     
 *   Random spacing                  Equal gaps between all shapes
 * 
 * How it works:
 * 1. Requires at least 3 shapes (2 edges + shapes to distribute between)
 * 2. Sorts shapes by their left position (left-to-right order)
 * 3. Keeps leftmost and rightmost shapes fixed as anchors
 * 4. Distributes middle shapes with equal gaps between them
 * 
 * Formula:
 * - totalSpace = rightmost.right - leftmost.left
 * - usedSpace = sum of all shape widths
 * - gapSpace = totalSpace - usedSpace
 * - gapSize = gapSpace / (numShapes - 1)
 * 
 * Note: If shapes don't fit in available space, they will overlap 
 * (same as PowerPoint behavior)
 * 
 * Use cases:
 * - Creating evenly spaced button rows
 * - Aligning timeline markers
 * - Organizing navigation elements
 * 
 * @returns {string} Status message
 */
function distributeHorizontally() {
  // Get the current selection from the presentation
  const selection = SlidesApp.getActivePresentation().getSelection();
  const pageElements = selection.getPageElementRange();

  // Validation: Must have at least 3 elements
  if (!pageElements) {
    return '⚠️ No elements selected. Please select at least 3 shapes.';
  }

  const elements = pageElements.getPageElements();
  
  if (elements.length < 3) {
    return `⚠️ Please select at least 3 shapes to distribute. Currently selected: ${elements.length}`;
  }

  // Sort shapes by left position (left to right)
  // We need to create a copy of the array to avoid modifying the original
  const sortedShapes = elements.slice().sort(function(a, b) {
    return a.getLeft() - b.getLeft();
  });

  // Calculate total space between leftmost and rightmost shapes
  const leftmostShape = sortedShapes[0];
  const rightmostShape = sortedShapes[sortedShapes.length - 1];
  
  // Left edge of the bounding area
  const leftEdge = leftmostShape.getLeft();
  
  // Right edge of the bounding area
  const rightEdge = rightmostShape.getLeft() + rightmostShape.getWidth();
  
  // Total horizontal space we're working with
  const totalSpace = rightEdge - leftEdge;

  // Calculate total width of all shapes
  let totalShapeWidth = 0;
  for (let i = 0; i < sortedShapes.length; i++) {
    totalShapeWidth += sortedShapes[i].getWidth();
  }

  // Calculate gap size
  // Total space minus all shape widths gives us space for gaps
  const totalGapSpace = totalSpace - totalShapeWidth;
  
  // Divide gap space by number of gaps (one less than number of shapes)
  const gapSize = totalGapSpace / (sortedShapes.length - 1);

  // Position each shape (leftmost stays fixed, rightmost will end up in correct position)
  let currentLeft = leftEdge;
  
  for (let i = 0; i < sortedShapes.length; i++) {
    const shape = sortedShapes[i];
    
    // Set position (first shape already at correct position, so skip it)
    if (i > 0) {
      shape.setLeft(currentLeft);
    }
    
    // Move current position for next shape
    // Add this shape's width plus the gap to get to the next shape's left edge
    currentLeft += shape.getWidth() + gapSize;
  }

  return `✅ ${elements.length} shapes distributed horizontally!`;
}

/**
 * distributeVertically()
 *
 * Distributes selected shapes evenly across vertical space.
 * Similar to PowerPoint's "Distribute Vertically" feature.
 * 
 * Visual example:
 * 
 *   BEFORE:                AFTER:
 *   
 *   [Shape 1]              [Shape 1]
 *   [Shape 2]                ↓ equal gap
 *                          [Shape 2]
 *   [Shape 3]       →        ↓ equal gap
 *          [Shape 4]       [Shape 3]
 *   [Shape 5]                ↓ equal gap
 *                          [Shape 4]
 *                            ↓ equal gap
 *                          [Shape 5]
 * 
 * How it works:
 * 1. Requires at least 3 shapes (2 edges + shapes to distribute between)
 * 2. Sorts shapes by their top position (top-to-bottom order)
 * 3. Keeps topmost and bottommost shapes fixed as anchors
 * 4. Distributes middle shapes with equal gaps between them
 * 
 * Formula:
 * - totalSpace = bottommost.bottom - topmost.top
 * - usedSpace = sum of all shape heights
 * - gapSpace = totalSpace - usedSpace
 * - gapSize = gapSpace / (numShapes - 1)
 * 
 * Note: If shapes don't fit in available space, they will overlap
 * (same as PowerPoint behavior)
 * 
 * Use cases:
 * - Creating evenly spaced vertical navigation
 * - Aligning list items
 * - Organizing vertical timelines
 * 
 * @returns {string} Status message
 */
function distributeVertically() {
  // Get the current selection from the presentation
  const selection = SlidesApp.getActivePresentation().getSelection();
  const pageElements = selection.getPageElementRange();

  // Validation: Must have at least 3 elements
  if (!pageElements) {
    return '⚠️ No elements selected. Please select at least 3 shapes.';
  }

  const elements = pageElements.getPageElements();
  
  if (elements.length < 3) {
    return `⚠️ Please select at least 3 shapes to distribute. Currently selected: ${elements.length}`;
  }

  // Sort shapes by top position (top to bottom)
  // We need to create a copy of the array to avoid modifying the original
  const sortedShapes = elements.slice().sort(function(a, b) {
    return a.getTop() - b.getTop();
  });

  // Calculate total space between topmost and bottommost shapes
  const topmostShape = sortedShapes[0];
  const bottommostShape = sortedShapes[sortedShapes.length - 1];
  
  // Top edge of the bounding area
  const topEdge = topmostShape.getTop();
  
  // Bottom edge of the bounding area
  const bottomEdge = bottommostShape.getTop() + bottommostShape.getHeight();
  
  // Total vertical space we're working with
  const totalSpace = bottomEdge - topEdge;

  // Calculate total height of all shapes
  let totalShapeHeight = 0;
  for (let i = 0; i < sortedShapes.length; i++) {
    totalShapeHeight += sortedShapes[i].getHeight();
  }

  // Calculate gap size
  // Total space minus all shape heights gives us space for gaps
  const totalGapSpace = totalSpace - totalShapeHeight;
  
  // Divide gap space by number of gaps (one less than number of shapes)
  const gapSize = totalGapSpace / (sortedShapes.length - 1);

  // Position each shape (topmost stays fixed, bottommost will end up in correct position)
  let currentTop = topEdge;
  
  for (let i = 0; i < sortedShapes.length; i++) {
    const shape = sortedShapes[i];
    
    // Set position (first shape already at correct position, so skip it)
    if (i > 0) {
      shape.setTop(currentTop);
    }
    
    // Move current position for next shape
    // Add this shape's height plus the gap to get to the next shape's top edge
    currentTop += shape.getHeight() + gapSize;
  }

  return `✅ ${elements.length} shapes distributed vertically!`;
}

/**
 * dockLeft()
 *
 * Moves selected shapes until their RIGHT edges touch the anchor's LEFT edge.
 * Creates a "docked to the left" effect where shapes stack to the left of the anchor.
 * 
 * Visual example:
 * 
 *   BEFORE:                         AFTER:
 *   
 *        [Shape 1]                 [Shape 1][Anchor]
 *   [Anchor]                  →    [Shape 2][Anchor]
 *      [Shape 2]
 *   
 *   Shapes dock to the left side of the anchor.
 * 
 * How it works:
 * 1. Resolves the anchor shape (using Set Anchor or fallback)
 * 2. For each non-anchor shape:
 *    - Calculates new position so shape.right = anchor.left
 *    - Moves shape horizontally (vertical position unchanged)
 * 
 * Formula: newLeft = anchorLeft - shapeWidth
 * 
 * Note: Multiple shapes will stack on top of each other at the same position.
 * To avoid overlap, consider docking shapes one at a time, or use this
 * intentionally to create layered effects.
 * 
 * Use cases:
 * - Creating side-by-side layouts
 * - Aligning labels to the left of content boxes
 * - Building flowcharts with connected shapes
 * 
 * @returns {string} Status message
 */
function dockLeft() {
  // Get the current selection from the presentation
  const selection = SlidesApp.getActivePresentation().getSelection();
  const pageElements = selection.getPageElementRange();

  // Validation: Must have page elements selected
  if (!pageElements) {
    return '⚠️ No elements selected. Please select at least 2 shapes.';
  }

  const elements = pageElements.getPageElements();
  
  // Validation: Need at least 2 elements (anchor + shapes to dock)
  if (elements.length < 2) {
    return '⚠️ Please select at least 2 shapes (anchor + shapes to dock).';
  }

  // Resolve anchor using the existing anchor resolution logic
  const anchor = getAnchorOrFallback(elements);
  const anchorLeft = anchor.getLeft();

  // Move each non-anchor shape
  let movedCount = 0;
  
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    
    // Skip the anchor itself - we don't want to move it
    if (el.getObjectId() === anchor.getObjectId()) {
      continue;
    }

    try {
      // Calculate new position: right edge touches anchor's left edge
      // Formula: newLeft = anchorLeft - shapeWidth
      // This positions the shape so: shapeLeft + shapeWidth = anchorLeft
      const newLeft = anchorLeft - el.getWidth();
      el.setLeft(newLeft);
      movedCount++;
    } catch (e) {
      // Skip shapes that can't be moved
    }
  }

  return `✅ ${movedCount} shape(s) docked to left of anchor!`;
}

/**
 * dockRight()
 *
 * Moves selected shapes until their LEFT edges touch the anchor's RIGHT edge.
 * Creates a "docked to the right" effect where shapes stack to the right of the anchor.
 * 
 * Visual example:
 * 
 *   BEFORE:                         AFTER:
 *   
 *   [Anchor]                       [Anchor][Shape 1]
 *        [Shape 1]            →    [Anchor][Shape 2]
 *   [Shape 2]
 *   
 *   Shapes dock to the right side of the anchor.
 * 
 * How it works:
 * 1. Resolves the anchor shape (using Set Anchor or fallback)
 * 2. For each non-anchor shape:
 *    - Calculates new position so shape.left = anchor.right
 *    - Moves shape horizontally (vertical position unchanged)
 * 
 * Formula: newLeft = anchorLeft + anchorWidth
 * 
 * Note: Multiple shapes will stack on top of each other at the same position.
 * To avoid overlap, consider docking shapes one at a time.
 * 
 * Use cases:
 * - Creating horizontal flowcharts
 * - Positioning labels to the right of icons
 * - Building timeline sequences
 * 
 * @returns {string} Status message
 */
function dockRight() {
  // Get the current selection from the presentation
  const selection = SlidesApp.getActivePresentation().getSelection();
  const pageElements = selection.getPageElementRange();

  // Validation: Must have page elements selected
  if (!pageElements) {
    return '⚠️ No elements selected. Please select at least 2 shapes.';
  }

  const elements = pageElements.getPageElements();
  
  // Validation: Need at least 2 elements (anchor + shapes to dock)
  if (elements.length < 2) {
    return '⚠️ Please select at least 2 shapes (anchor + shapes to dock).';
  }

  // Resolve anchor using the existing anchor resolution logic
  const anchor = getAnchorOrFallback(elements);
  
  // Calculate anchor's right edge position
  const anchorRight = anchor.getLeft() + anchor.getWidth();

  // Move each non-anchor shape
  let movedCount = 0;
  
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    
    // Skip the anchor itself - we don't want to move it
    if (el.getObjectId() === anchor.getObjectId()) {
      continue;
    }

    try {
      // Calculate new position: left edge touches anchor's right edge
      // Formula: newLeft = anchorRight
      // This positions the shape so: shapeLeft = anchorLeft + anchorWidth
      el.setLeft(anchorRight);
      movedCount++;
    } catch (e) {
      // Skip shapes that can't be moved
    }
  }

  return `✅ ${movedCount} shape(s) docked to right of anchor!`;
}

/**
 * dockTop()
 *
 * Moves selected shapes until their BOTTOM edges touch the anchor's TOP edge.
 * Creates a "docked above" effect where shapes stack above the anchor.
 * 
 * Visual example:
 * 
 *   BEFORE:                         AFTER:
 *   
 *   [Shape 1]                       [Shape 1]
 *                                   [Shape 2]
 *   [Anchor]                  →     [Anchor]
 *                     [Shape 2]
 *   
 *   Shapes dock above the anchor.
 * 
 * How it works:
 * 1. Resolves the anchor shape (using Set Anchor or fallback)
 * 2. For each non-anchor shape:
 *    - Calculates new position so shape.bottom = anchor.top
 *    - Moves shape vertically (horizontal position unchanged)
 * 
 * Formula: newTop = anchorTop - shapeHeight
 * 
 * Note: Multiple shapes will stack on top of each other at the same position.
 * To avoid overlap, consider docking shapes one at a time.
 * 
 * Use cases:
 * - Creating stacked vertical layouts
 * - Building vertical flowcharts
 * - Positioning headers above content sections
 * 
 * @returns {string} Status message
 */
function dockTop() {
  // Get the current selection from the presentation
  const selection = SlidesApp.getActivePresentation().getSelection();
  const pageElements = selection.getPageElementRange();

  // Validation: Must have page elements selected
  if (!pageElements) {
    return '⚠️ No elements selected. Please select at least 2 shapes.';
  }

  const elements = pageElements.getPageElements();
  
  // Validation: Need at least 2 elements (anchor + shapes to dock)
  if (elements.length < 2) {
    return '⚠️ Please select at least 2 shapes (anchor + shapes to dock).';
  }

  // Resolve anchor using the existing anchor resolution logic
  const anchor = getAnchorOrFallback(elements);
  const anchorTop = anchor.getTop();

  // Move each non-anchor shape
  let movedCount = 0;
  
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    
    // Skip the anchor itself - we don't want to move it
    if (el.getObjectId() === anchor.getObjectId()) {
      continue;
    }

    try {
      // Calculate new position: bottom edge touches anchor's top edge
      // Formula: newTop = anchorTop - shapeHeight
      // This positions the shape so: shapeTop + shapeHeight = anchorTop
      const newTop = anchorTop - el.getHeight();
      el.setTop(newTop);
      movedCount++;
    } catch (e) {
      // Skip shapes that can't be moved
    }
  }

  return `✅ ${movedCount} shape(s) docked above anchor!`;
}

/**
 * dockBottom()
 *
 * Moves selected shapes until their TOP edges touch the anchor's BOTTOM edge.
 * Creates a "docked below" effect where shapes stack below the anchor.
 * 
 * Visual example:
 * 
 *   BEFORE:                         AFTER:
 *   
 *   [Shape 1]                       [Anchor]
 *                                   [Shape 1]
 *   [Anchor]                  →     [Shape 2]
 *        [Shape 2]
 *   
 *   Shapes dock below the anchor.
 * 
 * How it works:
 * 1. Resolves the anchor shape (using Set Anchor or fallback)
 * 2. For each non-anchor shape:
 *    - Calculates new position so shape.top = anchor.bottom
 *    - Moves shape vertically (horizontal position unchanged)
 * 
 * Formula: newTop = anchorTop + anchorHeight
 * 
 * Note: Multiple shapes will stack on top of each other at the same position.
 * To avoid overlap, consider docking shapes one at a time.
 * 
 * Use cases:
 * - Creating dropdown-style layouts
 * - Positioning footers below content sections
 * - Building vertical step-by-step diagrams
 * 
 * @returns {string} Status message
 */
function dockBottom() {
  // Get the current selection from the presentation
  const selection = SlidesApp.getActivePresentation().getSelection();
  const pageElements = selection.getPageElementRange();

  // Validation: Must have page elements selected
  if (!pageElements) {
    return '⚠️ No elements selected. Please select at least 2 shapes.';
  }

  const elements = pageElements.getPageElements();
  
  // Validation: Need at least 2 elements (anchor + shapes to dock)
  if (elements.length < 2) {
    return '⚠️ Please select at least 2 shapes (anchor + shapes to dock).';
  }

  // Resolve anchor using the existing anchor resolution logic
  const anchor = getAnchorOrFallback(elements);
  
  // Calculate anchor's bottom edge position
  const anchorBottom = anchor.getTop() + anchor.getHeight();

  // Move each non-anchor shape
  let movedCount = 0;
  
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    
    // Skip the anchor itself - we don't want to move it
    if (el.getObjectId() === anchor.getObjectId()) {
      continue;
    }

    try {
      // Calculate new position: top edge touches anchor's bottom edge
      // Formula: newTop = anchorBottom
      // This positions the shape so: shapeTop = anchorTop + anchorHeight
      el.setTop(anchorBottom);
      movedCount++;
    } catch (e) {
      // Skip shapes that can't be moved
    }
  }

  return `✅ ${movedCount} shape(s) docked below anchor!`;
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

/**
 * getSelectedTableForSwap_()
 *
 * Resolves which table should be used for table swap operations.
 *
 * Resolution strategy:
 * 1. If the current selection includes exactly one table, use that table.
 * 2. If no table is selected, but the current slide has exactly one table,
 *    use that single table as a convenience fallback.
 * 3. Otherwise, return an explicit error message describing what the user
 *    needs to do (select one table, or reduce ambiguity).
 *
 * Why this helper exists:
 * - Keeps row and column swap functions focused only on swap logic
 * - Centralizes table-selection validation in one place
 * - Ensures both features behave consistently in edge cases
 *
 * @returns {Object} { table: Table|null, error: string|null }
 */
function getSelectedTableForSwap_() {
  // Access active presentation and current selection context.
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const range = selection.getPageElementRange();

  // Case 1: User selected page elements directly on slide canvas.
  if (range) {
    // Keep only TABLE elements from the selection.
    const tableElements = range.getPageElements().filter(function (el) {
      try {
        return el.getPageElementType() === SlidesApp.PageElementType.TABLE;
      } catch (e) {
        // Defensive fallback: skip elements that fail type resolution.
        return false;
      }
    });

    // Exactly one table selected -> unambiguous success.
    if (tableElements.length === 1) {
      return {table: tableElements[0].asTable(), error: null};
    }

    // More than one table selected -> ambiguous, ask for one table only.
    if (tableElements.length > 1) {
      return {table: null, error: 'Please select only one table.'};
    }
  }

  // Case 2: No selected table. Try slide-level fallback if safe to infer.
  const currentPage = selection.getCurrentPage();
  if (currentPage && currentPage.getPageType() === SlidesApp.PageType.SLIDE) {
    const tables = currentPage.asSlide().getTables();

    // Exactly one table on slide -> use it without forcing explicit selection.
    if (tables.length === 1) {
      return {table: tables[0], error: null};
    }
  }

  // No valid table context found.
  return {table: null, error: 'Please select a table first.'};
}

/**
 * validateTableIndex_(value, label, max)
 *
 * Validates a user-provided 1-based table index (row or column).
 *
 * Validation rules:
 * - Must be a whole number (integer)
 * - Must be within [1, max] where max is row/column count
 *
 * Returns both:
 * - `index`: normalized numeric value on success
 * - `error`: human-readable error message on failure
 *
 * @param {number} value - Raw index value from UI
 * @param {string} label - Human-readable label for error messages
 * @param {number} max - Maximum allowed value
 * @returns {Object} { index: number|null, error: string|null }
 */
function validateTableIndex_(value, label, max) {
  // Normalize input to Number for safe numeric checks.
  const index = Number(value);

  // Reject decimals, NaN, and other non-integer inputs.
  if (!Number.isInteger(index)) {
    return {index: null, error: `${label} must be a whole number.`};
  }

  // Enforce 1-based bounds expected by UI.
  if (index < 1 || index > max) {
    return {index: null, error: `${label} must be between 1 and ${max}.`};
  }

  // Validation passed.
  return {index: index, error: null};
}

/**
 * safeGet_(getter)
 *
 * Small utility wrapper for API calls that may throw if a style value
 * isn't explicitly set on a given text range/cell.
 *
 * @param {Function} getter - Function that returns a value
 * @returns {*} Returned value, or null if getter throws
 */
function safeGet_(getter) {
  try {
    return getter();
  } catch (e) {
    return null;
  }
}

/**
 * captureColorSpec_(color)
 *
 * Converts a Slides Color object into a lightweight spec that can be
 * reapplied later without retaining object references.
 *
 * @param {Color} color - Slides Color object
 * @returns {Object|null} { kind: 'rgb'|'theme', value: string|ThemeColorType }
 */
function captureColorSpec_(color) {
  if (!color) {
    return null;
  }

  const colorType = safeGet_(function () { return color.getColorType(); });
  if (colorType === SlidesApp.ColorType.RGB) {
    const rgb = safeGet_(function () { return color.asRgbColor(); });
    if (!rgb) {
      return null;
    }
    return {
      kind: 'rgb',
      value: safeGet_(function () { return rgb.asHexString(); })
    };
  }

  if (colorType === SlidesApp.ColorType.THEME) {
    const theme = safeGet_(function () { return color.asThemeColor(); });
    if (!theme) {
      return null;
    }
    return {
      kind: 'theme',
      value: safeGet_(function () { return theme.getThemeColorType(); })
    };
  }

  return null;
}

/**
 * applyTextColorSpec_(textStyle, mode, colorSpec)
 *
 * Applies captured color spec to TextStyle foreground/background.
 *
 * @param {TextStyle} textStyle - Target text style
 * @param {string} mode - 'foreground' or 'background'
 * @param {Object|null} colorSpec - Captured color spec
 */
function applyTextColorSpec_(textStyle, mode, colorSpec) {
  if (!colorSpec || !colorSpec.value) {
    return;
  }

  if (mode === 'foreground') {
    try { textStyle.setForegroundColor(colorSpec.value); } catch (e) {}
    return;
  }

  try { textStyle.setBackgroundColor(colorSpec.value); } catch (e) {}
}

/**
 * captureTextStyleSnapshot_(textStyle)
 *
 * Captures a portable snapshot of text style attributes for a run.
 * This snapshot is later re-applied to another range when
 * "keep formatting" is enabled.
 *
 * @param {TextStyle} textStyle - Slides TextStyle object
 * @returns {Object} Serializable style snapshot
 */
function captureTextStyleSnapshot_(textStyle) {
  return {
    bold: safeGet_(function () { return textStyle.isBold(); }),
    italic: safeGet_(function () { return textStyle.isItalic(); }),
    underline: safeGet_(function () { return textStyle.isUnderline(); }),
    strikethrough: safeGet_(function () { return textStyle.isStrikethrough(); }),
    smallCaps: safeGet_(function () { return textStyle.isSmallCaps(); }),
    fontFamily: safeGet_(function () { return textStyle.getFontFamily(); }),
    fontSize: safeGet_(function () { return textStyle.getFontSize(); }),
    fontWeight: safeGet_(function () { return textStyle.getFontWeight(); }),
    baselineOffset: safeGet_(function () { return textStyle.getBaselineOffset(); }),
    foregroundColor: captureColorSpec_(safeGet_(function () { return textStyle.getForegroundColor(); })),
    backgroundColor: captureColorSpec_(safeGet_(function () { return textStyle.getBackgroundColor(); })),
    isBackgroundTransparent: safeGet_(function () { return textStyle.isBackgroundTransparent(); }),
    linkUrl: safeGet_(function () {
      const link = textStyle.getLink();
      return link ? link.getUrl() : null;
    })
  };
}

/**
 * applyTextStyleSnapshot_(textStyle, snapshot)
 *
 * Applies a captured text-style snapshot onto a target TextStyle.
 * Each setter is isolated in try/catch to keep the swap operation resilient
 * when specific properties are unsupported on a given range.
 *
 * @param {TextStyle} textStyle - Target TextStyle object
 * @param {Object} snapshot - Style snapshot from captureTextStyleSnapshot_
 */
function applyTextStyleSnapshot_(textStyle, snapshot) {
  try {
    if (snapshot.linkUrl) {
      textStyle.setLinkUrl(snapshot.linkUrl);
    } else {
      textStyle.removeLink();
    }
  } catch (e) {}

  try { if (snapshot.fontFamily !== null && snapshot.fontWeight !== null) { textStyle.setFontFamilyAndWeight(snapshot.fontFamily, snapshot.fontWeight); } } catch (e) {}
  try { if (snapshot.fontFamily !== null) { textStyle.setFontFamily(snapshot.fontFamily); } } catch (e) {}
  try { if (snapshot.fontSize !== null) { textStyle.setFontSize(snapshot.fontSize); } } catch (e) {}
  try { if (snapshot.bold !== null) { textStyle.setBold(snapshot.bold); } } catch (e) {}
  try { if (snapshot.italic !== null) { textStyle.setItalic(snapshot.italic); } } catch (e) {}
  try { if (snapshot.underline !== null) { textStyle.setUnderline(snapshot.underline); } } catch (e) {}
  try { if (snapshot.strikethrough !== null) { textStyle.setStrikethrough(snapshot.strikethrough); } } catch (e) {}
  try { if (snapshot.smallCaps !== null) { textStyle.setSmallCaps(snapshot.smallCaps); } } catch (e) {}
  try { if (snapshot.baselineOffset !== null) { textStyle.setBaselineOffset(snapshot.baselineOffset); } } catch (e) {}
  applyTextColorSpec_(textStyle, 'foreground', snapshot.foregroundColor);

  if (snapshot.isBackgroundTransparent === true) {
    try { textStyle.setBackgroundColorTransparent(); } catch (e) {}
  } else if (snapshot.backgroundColor !== null) {
    applyTextColorSpec_(textStyle, 'background', snapshot.backgroundColor);
  }
}

/**
 * captureParagraphStyleSnapshot_(paragraphStyle)
 *
 * Captures paragraph-level styling so it can travel with content when
 * swapping rows/columns in "keep formatting" mode.
 *
 * @param {ParagraphStyle} paragraphStyle - Source paragraph style
 * @returns {Object} Paragraph style snapshot
 */
function captureParagraphStyleSnapshot_(paragraphStyle) {
  return {
    lineSpacing: safeGet_(function () { return paragraphStyle.getLineSpacing(); }),
    spaceAbove: safeGet_(function () { return paragraphStyle.getSpaceAbove(); }),
    spaceBelow: safeGet_(function () { return paragraphStyle.getSpaceBelow(); }),
    indentStart: safeGet_(function () { return paragraphStyle.getIndentStart(); }),
    indentEnd: safeGet_(function () { return paragraphStyle.getIndentEnd(); }),
    indentFirstLine: safeGet_(function () { return paragraphStyle.getIndentFirstLine(); }),
    spacingMode: safeGet_(function () { return paragraphStyle.getSpacingMode(); }),
    paragraphAlignment: safeGet_(function () { return paragraphStyle.getParagraphAlignment(); })
  };
}

/**
 * applyParagraphStyleSnapshot_(paragraphStyle, snapshot)
 *
 * Applies a paragraph style snapshot to target paragraph ranges.
 *
 * @param {ParagraphStyle} paragraphStyle - Target paragraph style
 * @param {Object} snapshot - Captured paragraph style data
 */
function applyParagraphStyleSnapshot_(paragraphStyle, snapshot) {
  try { if (snapshot.lineSpacing !== null) { paragraphStyle.setLineSpacing(snapshot.lineSpacing); } } catch (e) {}
  try { if (snapshot.spaceAbove !== null) { paragraphStyle.setSpaceAbove(snapshot.spaceAbove); } } catch (e) {}
  try { if (snapshot.spaceBelow !== null) { paragraphStyle.setSpaceBelow(snapshot.spaceBelow); } } catch (e) {}
  try { if (snapshot.indentStart !== null) { paragraphStyle.setIndentStart(snapshot.indentStart); } } catch (e) {}
  try { if (snapshot.indentEnd !== null) { paragraphStyle.setIndentEnd(snapshot.indentEnd); } } catch (e) {}
  try { if (snapshot.indentFirstLine !== null) { paragraphStyle.setIndentFirstLine(snapshot.indentFirstLine); } } catch (e) {}
  try { if (snapshot.spacingMode !== null) { paragraphStyle.setSpacingMode(snapshot.spacingMode); } } catch (e) {}
  try { if (snapshot.paragraphAlignment !== null) { paragraphStyle.setParagraphAlignment(snapshot.paragraphAlignment); } } catch (e) {}
}

/**
 * captureCellPayload_(cell)
 *
 * Captures all swap-relevant data from a table cell:
 * - text content
 * - per-run text style
 * - per-paragraph style
 * - cell content alignment
 * - fill state/color
 *
 * @param {TableCell} cell - Source table cell
 * @returns {Object} Captured payload
 */
function captureCellPayload_(cell) {
  const text = cell.getText();
  const payload = {
    plainText: text.asString(),
    contentAlignment: safeGet_(function () { return cell.getContentAlignment(); }),
    fillVisible: null,
    fillType: null,
    fillColor: null,
    fillAlpha: null,
    runs: [],
    paragraphs: []
  };

  const fill = safeGet_(function () { return cell.getFill(); });
  if (fill) {
    payload.fillVisible = safeGet_(function () { return fill.isVisible(); });
    payload.fillType = safeGet_(function () { return fill.getType(); });

    const solid = safeGet_(function () { return fill.getSolidFill(); });
    if (solid) {
      payload.fillColor = captureColorSpec_(safeGet_(function () { return solid.getColor(); }));
      payload.fillAlpha = safeGet_(function () { return solid.getAlpha(); });
    }
  }

  const runs = safeGet_(function () { return text.getRuns(); }) || [];
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    payload.runs.push({
      start: safeGet_(function () { return run.getStartIndex(); }),
      end: safeGet_(function () { return run.getEndIndex(); }),
      style: captureTextStyleSnapshot_(run.getTextStyle())
    });
  }

  const paragraphs = safeGet_(function () { return text.getParagraphs(); }) || [];
  for (let p = 0; p < paragraphs.length; p++) {
    const paragraph = paragraphs[p];
    const range = safeGet_(function () { return paragraph.getRange(); });
    if (!range) {
      continue;
    }

    payload.paragraphs.push({
      start: safeGet_(function () { return range.getStartIndex(); }),
      end: safeGet_(function () { return range.getEndIndex(); }),
      style: captureParagraphStyleSnapshot_(range.getParagraphStyle())
    });
  }

  return payload;
}

/**
 * applyCellPayload_(cell, payload)
 *
 * Applies a captured payload onto a target table cell.
 * Operation order:
 * 1. Set plain text baseline
 * 2. Apply fill and cell-level alignment
 * 3. Reapply paragraph styles
 * 4. Reapply run-level text styles
 *
 * @param {TableCell} cell - Destination cell
 * @param {Object} payload - Captured payload from captureCellPayload_
 */
function applyCellPayload_(cell, payload) {
  const text = cell.getText();
  text.setText(payload.plainText || '');

  if (payload.contentAlignment !== null) {
    try { cell.setContentAlignment(payload.contentAlignment); } catch (e) {}
  }

  const fill = safeGet_(function () { return cell.getFill(); });
  if (fill && payload.fillVisible === false) {
    try { fill.setTransparent(); } catch (e) {}
  } else if (fill && payload.fillType === SlidesApp.FillType.SOLID && payload.fillColor && payload.fillColor.value) {
    if (payload.fillAlpha !== null) {
      try {
        fill.setSolidFill(payload.fillColor.value, payload.fillAlpha);
      } catch (e) {
        try { fill.setSolidFill(payload.fillColor.value); } catch (e2) {}
      }
    } else {
      try { fill.setSolidFill(payload.fillColor.value); } catch (e) {}
    }
  }

  const textLength = safeGet_(function () { return text.asString().length; }) || 0;
  for (let p = 0; p < payload.paragraphs.length; p++) {
    const paragraph = payload.paragraphs[p];
    const start = Math.max(0, Math.min(textLength, Number(paragraph.start)));
    const end = Math.max(0, Math.min(textLength, Number(paragraph.end)));

    if (start >= end) {
      continue;
    }

    const range = safeGet_(function () { return text.getRange(start, end); });
    if (!range) {
      continue;
    }

    applyParagraphStyleSnapshot_(range.getParagraphStyle(), paragraph.style);
  }

  for (let r = 0; r < payload.runs.length; r++) {
    const run = payload.runs[r];
    const start = Math.max(0, Math.min(textLength, Number(run.start)));
    const end = Math.max(0, Math.min(textLength, Number(run.end)));

    if (start >= end) {
      continue;
    }

    const range = safeGet_(function () { return text.getRange(start, end); });
    if (!range) {
      continue;
    }

    applyTextStyleSnapshot_(range.getTextStyle(), run.style);
  }
}

/**
 * isMergedCell_(cell)
 *
 * Checks whether a cell is in a merged state. Swapping merged cells can produce
 * ambiguous behavior, so we block those operations and return a clear message.
 *
 * @param {TableCell} cell - Target cell
 * @returns {boolean} True if merged, false otherwise
 */
function isMergedCell_(cell) {
  const mergeState = safeGet_(function () { return cell.getMergeState(); });

  // If merge state cannot be read, treat as non-merged so we don't block swaps.
  if (!mergeState) {
    return false;
  }

  // Avoid direct enum dependency (it can vary in Apps Script surfaces).
  // We normalize by string value and only treat explicit "NORMAL" as unmerged.
  const stateString = String(mergeState).toUpperCase();
  return stateString !== 'NORMAL';
}

/**
 * swapCells_(cellA, cellB, keepFormatting)
 *
 * Swaps two cells either as plain text (default) or with style payload.
 *
 * @param {TableCell} cellA - First cell
 * @param {TableCell} cellB - Second cell
 * @param {boolean} keepFormatting - Whether style/fill should move with text
 * @returns {Object} { ok: boolean, error: string|null }
 */
function swapCells_(cellA, cellB, keepFormatting) {
  if (isMergedCell_(cellA) || isMergedCell_(cellB)) {
    return {
      ok: false,
      error: 'Swap with merged cells is not supported yet. Please unmerge cells first.'
    };
  }

  if (!keepFormatting) {
    const textA = cellA.getText().asString();
    const textB = cellB.getText().asString();
    cellA.getText().setText(textB);
    cellB.getText().setText(textA);
    return {ok: true, error: null};
  }

  const payloadA = captureCellPayload_(cellA);
  const payloadB = captureCellPayload_(cellB);
  applyCellPayload_(cellA, payloadB);
  applyCellPayload_(cellB, payloadA);

  return {ok: true, error: null};
}

/**
 * swapTableRows(rowA, rowB, keepFormatting)
 *
 * Swaps row contents in a selected table.
 *
 * Modes:
 * - keepFormatting = false: swap plain text only (current behavior)
 * - keepFormatting = true: swap text + text styles + paragraph styles +
 *   cell fill + cell content alignment
 *
 * Note about borders:
 * - Google Slides Apps Script currently does not expose table-cell border
 *   setters/getters in the same way as text/fill, so borders cannot be
 *   programmatically swapped in this implementation.
 *
 * @param {number} rowA - First row number (1-based)
 * @param {number} rowB - Second row number (1-based)
 * @param {boolean} keepFormatting - Preserve source formatting with content
 * @returns {string} Status message
 */
function swapTableRows(rowA, rowB, keepFormatting) {
  const tableResult = getSelectedTableForSwap_();
  if (tableResult.error) {
    return tableResult.error;
  }

  const table = tableResult.table;
  const numRows = table.getNumRows();
  const numCols = table.getNumColumns();
  const preserveFormatting = keepFormatting === true;

  const first = validateTableIndex_(rowA, 'First row', numRows);
  if (first.error) {
    return first.error;
  }

  const second = validateTableIndex_(rowB, 'Second row', numRows);
  if (second.error) {
    return second.error;
  }

  if (first.index === second.index) {
    return 'Row numbers are identical. Nothing to swap.';
  }

  const rowIndexA = first.index - 1;
  const rowIndexB = second.index - 1;

  for (let col = 0; col < numCols; col++) {
    const cellA = table.getCell(rowIndexA, col);
    const cellB = table.getCell(rowIndexB, col);
    const result = swapCells_(cellA, cellB, preserveFormatting);

    if (!result.ok) {
      return result.error;
    }
  }

  if (preserveFormatting) {
    return `Swapped row ${first.index} with row ${second.index} with source text/cell formatting. Border swap is not supported by Apps Script.`;
  }
  return `Swapped row ${first.index} with row ${second.index}.`;
}

/**
 * swapTableColumns(colA, colB, keepFormatting)
 *
 * Swaps column contents in a selected table.
 *
 * Modes:
 * - keepFormatting = false: swap plain text only (current behavior)
 * - keepFormatting = true: swap text + text styles + paragraph styles +
 *   cell fill + cell content alignment
 *
 * Note about borders:
 * - Google Slides Apps Script currently does not expose table-cell border
 *   setters/getters in the same way as text/fill, so borders cannot be
 *   programmatically swapped in this implementation.
 *
 * @param {number} colA - First column number (1-based)
 * @param {number} colB - Second column number (1-based)
 * @param {boolean} keepFormatting - Preserve source formatting with content
 * @returns {string} Status message
 */
function swapTableColumns(colA, colB, keepFormatting) {
  const tableResult = getSelectedTableForSwap_();
  if (tableResult.error) {
    return tableResult.error;
  }

  const table = tableResult.table;
  const numRows = table.getNumRows();
  const numCols = table.getNumColumns();
  const preserveFormatting = keepFormatting === true;

  const first = validateTableIndex_(colA, 'First column', numCols);
  if (first.error) {
    return first.error;
  }

  const second = validateTableIndex_(colB, 'Second column', numCols);
  if (second.error) {
    return second.error;
  }

  if (first.index === second.index) {
    return 'Column numbers are identical. Nothing to swap.';
  }

  const colIndexA = first.index - 1;
  const colIndexB = second.index - 1;

  for (let row = 0; row < numRows; row++) {
    const cellA = table.getCell(row, colIndexA);
    const cellB = table.getCell(row, colIndexB);
    const result = swapCells_(cellA, cellB, preserveFormatting);

    if (!result.ok) {
      return result.error;
    }
  }

  if (preserveFormatting) {
    return `Swapped column ${first.index} with column ${second.index} with source text/cell formatting. Border swap is not supported by Apps Script.`;
  }
  return `Swapped column ${first.index} with column ${second.index}.`;
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


// ============================================================================
// SLIDE LAYOUT
// ============================================================================
//
// These functions insert pre-formatted text boxes to create common slide layouts.
// 
// Features:
// - Insert 2, 3, or 4 column layouts with title and content boxes
// - Insert footnote bars at the bottom of slides
// - All text boxes are properly sized and positioned
// - Text is pre-formatted with appropriate sizes and styling
//
// Layout specifications:
// - Vertical start: 0.75 inches from top of slide
// - Horizontal range: 0 to 9.1 inches
// - Title boxes: 1 inch tall, 16pt bold text
// - Content boxes: maximize remaining space, 14pt text
// - Gap between title and content: 0.05 inches
// - Gap between columns: 0.2 inches
//
// ============================================================================

/**
 * insertTwoColumns()
 *
 * Inserts a 2-column layout with title and content text boxes.
 * 
 * Layout:
 * - 2 columns, each with:
 *   - Title box: "Header 1", "Header 2" (1 inch tall, bold, 16pt)
 *   - Content box: "Content 1", "Content 2" (3.0 inches tall, 14pt)
 * - Positioned 0.75 inches from top of slide
 * - Spread across 0 to 9.1 inches horizontal space
 * - 0.2 inch gap between columns
 * - 0.05 inch gap between title and content boxes
 * 
 * Visual example:
 * 
 *   ┌──────────────┐  ┌──────────────┐
 *   │  Header 1    │  │  Header 2    │  ← Title boxes (1")
 *   ├──────────────┤  ├──────────────┤
 *   │              │  │              │
 *   │  Content 1   │  │  Content 2   │  ← Content boxes (3")
 *   │              │  │              │
 *   └──────────────┘  └──────────────┘
 * 
 * Use cases:
 * - Side-by-side comparisons
 * - Two-topic presentations
 * - Before/after layouts
 * - Feature comparison slides
 * 
 * @returns {string} Status message
 */
function insertTwoColumns() {
  // Get the current slide
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const currentPage = selection.getCurrentPage();
  
  // Validate that we're on a slide (not master or layout)
  if (!currentPage || currentPage.getPageType() !== SlidesApp.PageType.SLIDE) {
    return '⚠️ Please select a slide first. Click on a slide in the main canvas.';
  }
  
  const slide = currentPage.asSlide();
  
  // Layout constants (converting inches to points: 1 inch = 72 points)
  const leftMargin = 0.33 * 72;              // 23.76 points from left edge
  const verticalStart = 1.4 * 72;            // 100.8 points from top of slide (shifted down)
  const titleHeight = 0.6 * 72;              // 43.2 points tall for title boxes (shorter headers)
  const contentHeight = 2.95 * 72;           // 212.4 points tall (ends at 5.0" to leave room for footnote)
  const gapBetweenTitleAndContent = 0.05 * 72; // 3.6 points gap between title and content
  const columnGap = 0.2 * 72;                // 14.4 points gap between columns
  const numColumns = 2;                      // Number of columns to create
  const totalWidth = 9.34 * 72;              // 672.48 points total horizontal space (10" - 0.33" - 0.33")
  const padding = 0.01 * 72;                 // 0.72 points internal padding
  
  // Calculate column width
  // Formula: (total width - gaps between columns) / number of columns
  const columnWidth = (totalWidth - (numColumns - 1) * columnGap) / numColumns;
  
  // Create boxes for each column
  for (let i = 0; i < numColumns; i++) {
    // Calculate left position for this column (starting from left margin)
    const columnLeft = leftMargin + (i * (columnWidth + columnGap));
    
    // Create title box
    const titleBox = slide.insertTextBox(
      `Header ${i + 1}`,  // Text content
      columnLeft,         // Left position (X)
      verticalStart,      // Top position (Y)
      columnWidth,        // Width
      titleHeight         // Height
    );
    
    // Format title text: bold, 16pt, with padding
    const titleTextRange = titleBox.getText();
    titleTextRange.getTextStyle()
      .setBold(true)
      .setFontSize(16);
    
    // Add internal padding to title box
    const titleParagraphStyle = titleTextRange.getParagraphStyle();
    titleParagraphStyle
      .setIndentStart(padding)
      .setIndentEnd(padding)
      .setSpaceAbove(padding)
      .setSpaceBelow(padding);
    
    // Create content box below title (with small gap)
    const contentTop = verticalStart + titleHeight + gapBetweenTitleAndContent;
    const contentBox = slide.insertTextBox(
      `Content ${i + 1}`,  // Text content
      columnLeft,          // Left position (X)
      contentTop,          // Top position (Y)
      columnWidth,         // Width
      contentHeight        // Height
    );
    
    // Format content text: 14pt, with padding
    const contentTextRange = contentBox.getText();
    contentTextRange.getTextStyle()
      .setFontSize(14);
    
    // Add internal padding to content box
    const contentParagraphStyle = contentTextRange.getParagraphStyle();
    contentParagraphStyle
      .setIndentStart(padding)
      .setIndentEnd(padding)
      .setSpaceAbove(padding)
      .setSpaceBelow(padding);
  }
  
  return `✅ Inserted 2-column layout (${numColumns} titles + ${numColumns} content boxes)!`;
}

/**
 * insertThreeColumns()
 *
 * Inserts a 3-column layout with title and content text boxes.
 * 
 * Layout:
 * - 3 columns, each with:
 *   - Title box: "Header 1", "Header 2", "Header 3" (1 inch tall, bold, 16pt)
 *   - Content box: "Content 1", "Content 2", "Content 3" (3.0 inches tall, 14pt)
 * - Positioned 0.75 inches from top of slide
 * - Spread across 0 to 9.1 inches horizontal space
 * - 0.2 inch gap between columns
 * - 0.05 inch gap between title and content boxes
 * 
 * Visual example:
 * 
 *   ┌─────────┐  ┌─────────┐  ┌─────────┐
 *   │Header 1 │  │Header 2 │  │Header 3 │  ← Title boxes (1")
 *   ├─────────┤  ├─────────┤  ├─────────┤
 *   │         │  │         │  │         │
 *   │Content 1│  │Content 2│  │Content 3│  ← Content boxes (3")
 *   │         │  │         │  │         │
 *   └─────────┘  └─────────┘  └─────────┘
 * 
 * Use cases:
 * - Three-step processes
 * - Past/Present/Future comparisons
 * - Three-feature showcases
 * - Triple comparison slides
 * 
 * @returns {string} Status message
 */
function insertThreeColumns() {
  // Get the current slide
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const currentPage = selection.getCurrentPage();
  
  // Validate that we're on a slide (not master or layout)
  if (!currentPage || currentPage.getPageType() !== SlidesApp.PageType.SLIDE) {
    return '⚠️ Please select a slide first. Click on a slide in the main canvas.';
  }
  
  const slide = currentPage.asSlide();
  
  // Layout constants (converting inches to points: 1 inch = 72 points)
  const leftMargin = 0.33 * 72;              // 23.76 points from left edge
  const verticalStart = 1.4 * 72;            // 100.8 points from top of slide (shifted down)
  const titleHeight = 0.6 * 72;              // 43.2 points tall for title boxes (shorter headers)
  const contentHeight = 2.95 * 72;           // 212.4 points tall (ends at 5.0" to leave room for footnote)
  const gapBetweenTitleAndContent = 0.05 * 72; // 3.6 points gap between title and content
  const columnGap = 0.2 * 72;                // 14.4 points gap between columns
  const numColumns = 3;                      // Number of columns to create
  const totalWidth = 9.34 * 72;              // 672.48 points total horizontal space (10" - 0.33" - 0.33")
  const padding = 0.01 * 72;                 // 0.72 points internal padding
  
  // Calculate column width
  // Formula: (total width - gaps between columns) / number of columns
  const columnWidth = (totalWidth - (numColumns - 1) * columnGap) / numColumns;
  
  // Create boxes for each column
  for (let i = 0; i < numColumns; i++) {
    // Calculate left position for this column (starting from left margin)
    const columnLeft = leftMargin + (i * (columnWidth + columnGap));
    
    // Create title box
    const titleBox = slide.insertTextBox(
      `Header ${i + 1}`,  // Text content
      columnLeft,         // Left position (X)
      verticalStart,      // Top position (Y)
      columnWidth,        // Width
      titleHeight         // Height
    );
    
    // Format title text: bold, 16pt, with padding
    const titleTextRange = titleBox.getText();
    titleTextRange.getTextStyle()
      .setBold(true)
      .setFontSize(16);
    
    // Add internal padding to title box
    const titleParagraphStyle = titleTextRange.getParagraphStyle();
    titleParagraphStyle
      .setIndentStart(padding)
      .setIndentEnd(padding)
      .setSpaceAbove(padding)
      .setSpaceBelow(padding);
    
    // Create content box below title (with small gap)
    const contentTop = verticalStart + titleHeight + gapBetweenTitleAndContent;
    const contentBox = slide.insertTextBox(
      `Content ${i + 1}`,  // Text content
      columnLeft,          // Left position (X)
      contentTop,          // Top position (Y)
      columnWidth,         // Width
      contentHeight        // Height
    );
    
    // Format content text: 14pt, with padding
    const contentTextRange = contentBox.getText();
    contentTextRange.getTextStyle()
      .setFontSize(14);
    
    // Add internal padding to content box
    const contentParagraphStyle = contentTextRange.getParagraphStyle();
    contentParagraphStyle
      .setIndentStart(padding)
      .setIndentEnd(padding)
      .setSpaceAbove(padding)
      .setSpaceBelow(padding);
  }
  
  return `✅ Inserted 3-column layout (${numColumns} titles + ${numColumns} content boxes)!`;
}

/**
 * insertFourColumns()
 *
 * Inserts a 4-column layout with title and content text boxes.
 * 
 * Layout:
 * - 4 columns, each with:
 *   - Title box: "Header 1" through "Header 4" (1 inch tall, bold, 16pt)
 *   - Content box: "Content 1" through "Content 4" (3.0 inches tall, 14pt)
 * - Positioned 0.75 inches from top of slide
 * - Spread across 0 to 9.1 inches horizontal space
 * - 0.2 inch gap between columns
 * - 0.05 inch gap between title and content boxes
 * 
 * Visual example:
 * 
 *   ┌──────┐  ┌──────┐  ┌──────┐  ┌──────┐
 *   │Head 1│  │Head 2│  │Head 3│  │Head 4│  ← Title boxes (1")
 *   ├──────┤  ├──────┤  ├──────┤  ├──────┤
 *   │      │  │      │  │      │  │      │
 *   │Cont 1│  │Cont 2│  │Cont 3│  │Cont 4│  ← Content boxes (3")
 *   │      │  │      │  │      │  │      │
 *   └──────┘  └──────┘  └──────┘  └──────┘
 * 
 * Use cases:
 * - Four-step processes or workflows
 * - Quarterly comparisons
 * - Four-feature showcases
 * - Multi-option comparisons
 * 
 * Note: With 4 columns, each column is narrower (~2.125 inches).
 * Consider using shorter text or reducing font size if content doesn't fit.
 * 
 * @returns {string} Status message
 */
function insertFourColumns() {
  // Get the current slide
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const currentPage = selection.getCurrentPage();
  
  // Validate that we're on a slide (not master or layout)
  if (!currentPage || currentPage.getPageType() !== SlidesApp.PageType.SLIDE) {
    return '⚠️ Please select a slide first. Click on a slide in the main canvas.';
  }
  
  const slide = currentPage.asSlide();
  
  // Layout constants (converting inches to points: 1 inch = 72 points)
  const leftMargin = 0.33 * 72;              // 23.76 points from left edge
  const verticalStart = 1.4 * 72;            // 100.8 points from top of slide (shifted down)
  const titleHeight = 0.6 * 72;              // 43.2 points tall for title boxes (shorter headers)
  const contentHeight = 2.95 * 72;           // 212.4 points tall (ends at 5.0" to leave room for footnote)
  const gapBetweenTitleAndContent = 0.05 * 72; // 3.6 points gap between title and content
  const columnGap = 0.2 * 72;                // 14.4 points gap between columns
  const numColumns = 4;                      // Number of columns to create
  const totalWidth = 9.34 * 72;              // 672.48 points total horizontal space (10" - 0.33" - 0.33")
  const padding = 0.01 * 72;                 // 0.72 points internal padding
  
  // Calculate column width
  // Formula: (total width - gaps between columns) / number of columns
  const columnWidth = (totalWidth - (numColumns - 1) * columnGap) / numColumns;
  
  // Create boxes for each column
  for (let i = 0; i < numColumns; i++) {
    // Calculate left position for this column (starting from left margin)
    const columnLeft = leftMargin + (i * (columnWidth + columnGap));
    
    // Create title box
    const titleBox = slide.insertTextBox(
      `Header ${i + 1}`,  // Text content
      columnLeft,         // Left position (X)
      verticalStart,      // Top position (Y)
      columnWidth,        // Width
      titleHeight         // Height
    );
    
    // Format title text: bold, 16pt, with padding
    const titleTextRange = titleBox.getText();
    titleTextRange.getTextStyle()
      .setBold(true)
      .setFontSize(16);
    
    // Add internal padding to title box
    const titleParagraphStyle = titleTextRange.getParagraphStyle();
    titleParagraphStyle
      .setIndentStart(padding)
      .setIndentEnd(padding)
      .setSpaceAbove(padding)
      .setSpaceBelow(padding);
    
    // Create content box below title (with small gap)
    const contentTop = verticalStart + titleHeight + gapBetweenTitleAndContent;
    const contentBox = slide.insertTextBox(
      `Content ${i + 1}`,  // Text content
      columnLeft,          // Left position (X)
      contentTop,          // Top position (Y)
      columnWidth,         // Width
      contentHeight        // Height
    );
    
    // Format content text: 14pt, with padding
    const contentTextRange = contentBox.getText();
    contentTextRange.getTextStyle()
      .setFontSize(14);
    
    // Add internal padding to content box
    const contentParagraphStyle = contentTextRange.getParagraphStyle();
    contentParagraphStyle
      .setIndentStart(padding)
      .setIndentEnd(padding)
      .setSpaceAbove(padding)
      .setSpaceBelow(padding);
  }
  
  return `✅ Inserted 4-column layout (${numColumns} titles + ${numColumns} content boxes)!`;
}

/**
 * insertFootnote()
 *
 * Inserts a footnote text box at the bottom of the slide.
 * 
 * Layout:
 * - Position: 4.8 inches from top, 0 inches from left
 * - Width: 9 inches
 * - Height: 0.5 inches (auto-expands if content is longer)
 * - Text: "Footnote: 1. commentary 1, 2. commentary 2, 3. commentary 3"
 * - Font size: 12pt
 * 
 * Visual example:
 * 
 *   [Main slide content above]
 *   
 *   ─────────────────────────────────────────────
 *   Footnote: 1. commentary 1, 2. commentary 2, 3. commentary 3
 *   
 * Use cases:
 * - Adding citations or references
 * - Including disclaimers or legal text
 * - Providing additional context
 * - Sourcing data or statistics
 * 
 * Note: The footnote text is fully editable after insertion.
 * You can replace the placeholder commentary with your actual footnotes.
 * 
 * @returns {string} Status message
 */
function insertFootnote() {
  // Get the current slide
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const currentPage = selection.getCurrentPage();
  
  // Validate that we're on a slide (not master or layout)
  if (!currentPage || currentPage.getPageType() !== SlidesApp.PageType.SLIDE) {
    return '⚠️ Please select a slide first. Click on a slide in the main canvas.';
  }
  
  const slide = currentPage.asSlide();
  
  // Footnote layout constants (converting inches to points: 1 inch = 72 points)
  // Google Slides standard dimensions: 10" wide x 5.625" tall (720 x 405 points)
  const footnoteTop = 5.05 * 72;      // 363.6 points from top
  const footnoteLeft = 0.33 * 72;     // 23.76 points from left edge (matches column margin)
  const footnoteWidth = 9.34 * 72;    // 672.48 points wide (10" - 0.33" - 0.33" = 9.34")
  const footnoteHeight = 0.5 * 72;    // 36 points tall (will auto-expand if needed)
  
  // Internal padding (0.01 inches = 0.72 points)
  const padding = 0.01 * 72;          // 0.72 points padding
  
  // Create footnote text box
  const footnoteBox = slide.insertTextBox(
    'Footnotes: 1. Footnotes use 9-point text 2. Items should be numbered and presented without commas.',
    footnoteLeft,
    footnoteTop,
    footnoteWidth,
    footnoteHeight
  );

  // Format footnote text: 9pt with internal padding and resize to fit text
  const textRange = footnoteBox.getText();
  textRange.getTextStyle()
    .setFontSize(9);
  
  // Add internal padding to the text box
  const paragraphStyle = textRange.getParagraphStyle();
  paragraphStyle
    .setIndentStart(padding)        // Left padding
    .setIndentEnd(padding)          // Right padding
    .setSpaceAbove(padding)         // Top padding
    .setSpaceBelow(padding);        // Bottom padding
  
  // Note: Google Apps Script doesn't fully support programmatic autofit settings
  // Users may need to manually enable "Resize shape to fit text" from Format Options
  // if they want the box to automatically resize as they add more text
  
  return '✅ Footnote inserted at bottom of slide!';
}

// ============================================================================
// SIZE MANIPULATION
// ============================================================================

/**
 * Align Width - Makes all selected objects have the same width as the anchor
 * 
 * How it works:
 * 1. Identifies the anchor object (using Set Anchor or fallback to last element)
 * 2. Gets the anchor's width
 * 3. Sets all other selected objects to that same width (keeps their position)
 * 
 * Example:
 * - Anchor width: 200 points
 * - Object A width: 150 points → becomes 200 points
 * - Object B width: 300 points → becomes 200 points
 * 
 * Use cases:
 * - Creating uniform button sizes
 * - Making all cards the same width
 * - Standardizing column widths
 * 
 * @returns {string} Status message
 */
function alignWidth() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to align width';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const targetWidth = anchor.getWidth();
  
  // Align width of all non-anchor elements
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      element.setWidth(targetWidth);
      successCount++;
    } catch (e) {
      // Skip elements that don't support setWidth
    }
  }
  
  return `✅ Aligned width of ${successCount} object(s) to anchor (${Math.round(targetWidth / 72 * 100) / 100}")`;
}

/**
 * Align Height - Makes all selected objects have the same height as the anchor
 * 
 * How it works:
 * 1. Identifies the anchor object (using Set Anchor or fallback to last element)
 * 2. Gets the anchor's height
 * 3. Sets all other selected objects to that same height (keeps their position)
 * 
 * Example:
 * - Anchor height: 150 points
 * - Object A height: 100 points → becomes 150 points
 * - Object B height: 200 points → becomes 150 points
 * 
 * Use cases:
 * - Creating uniform row heights
 * - Making all icons the same height
 * - Standardizing image heights
 * 
 * @returns {string} Status message
 */
function alignHeight() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to align height';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const targetHeight = anchor.getHeight();
  
  // Align height of all non-anchor elements
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      element.setHeight(targetHeight);
      successCount++;
    } catch (e) {
      // Skip elements that don't support setHeight
    }
  }
  
  return `✅ Aligned height of ${successCount} object(s) to anchor (${Math.round(targetHeight / 72 * 100) / 100}")`;
}

/**
 * Align Both - Makes all selected objects have the same width AND height as the anchor
 * 
 * How it works:
 * 1. Identifies the anchor object (using Set Anchor or fallback to last element)
 * 2. Gets the anchor's width and height
 * 3. Sets all other selected objects to that same width and height (keeps their position)
 * 
 * Example:
 * - Anchor: 200 x 150 points
 * - Object A: 180 x 100 → becomes 200 x 150
 * - Object B: 250 x 200 → becomes 200 x 150
 * 
 * Use cases:
 * - Creating perfectly uniform elements
 * - Making all icons exactly the same size
 * - Standardizing photo dimensions
 * 
 * @returns {string} Status message
 */
function alignBoth() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to align size';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const targetWidth = anchor.getWidth();
  const targetHeight = anchor.getHeight();
  
  // Align both dimensions of all non-anchor elements
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      element.setWidth(targetWidth);
      element.setHeight(targetHeight);
      successCount++;
    } catch (e) {
      // Skip elements that don't support setWidth/setHeight
    }
  }
  
  const widthInches = Math.round(targetWidth / 72 * 100) / 100;
  const heightInches = Math.round(targetHeight / 72 * 100) / 100;
  return `✅ Aligned size of ${successCount} object(s) to anchor (${widthInches}" × ${heightInches}")`;
}

// ============================================================================
// STRETCH OBJECTS
// ============================================================================

/**
 * Stretch Left - Extends objects' left edge to match anchor's left edge
 * 
 * How it works:
 * 1. Identifies the anchor object
 * 2. For each selected object, extends its left edge to align with anchor's left edge
 * 3. Right edge stays fixed (object grows/shrinks leftward)
 * 
 * Example:
 * - Anchor left edge: 100 points
 * - Object left edge: 150 points, right edge: 300 points
 * - After: left edge: 100 points, right edge: 300 points (width increased by 50)
 * 
 * Use cases:
 * - Extending objects to a common left boundary
 * - Creating left-aligned blocks with consistent left edge
 * 
 * @returns {string} Status message
 */
function stretchLeft() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to stretch';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const targetLeft = anchor.getLeft();
  
  // Stretch left edge of all non-anchor elements
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      const currentLeft = element.getLeft();
      const currentRight = currentLeft + element.getWidth();
      const newWidth = currentRight - targetLeft;
      
      if (newWidth > 0) {
        element.setLeft(targetLeft);
        element.setWidth(newWidth);
        successCount++;
      }
    } catch (e) {
      // Skip elements that don't support these operations
    }
  }
  
  return `✅ Stretched left edge of ${successCount} object(s) to anchor`;
}

/**
 * Stretch Right - Extends objects' right edge to match anchor's right edge
 * 
 * How it works:
 * 1. Identifies the anchor object
 * 2. For each selected object, extends its right edge to align with anchor's right edge
 * 3. Left edge stays fixed (object grows/shrinks rightward)
 * 
 * Example:
 * - Anchor right edge: 400 points
 * - Object left edge: 100 points, right edge: 300 points
 * - After: left edge: 100 points, right edge: 400 points (width increased by 100)
 * 
 * Use cases:
 * - Extending objects to a common right boundary
 * - Creating right-aligned blocks with consistent right edge
 * 
 * @returns {string} Status message
 */
function stretchRight() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to stretch';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const targetRight = anchor.getLeft() + anchor.getWidth();
  
  // Stretch right edge of all non-anchor elements
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      const currentLeft = element.getLeft();
      const newWidth = targetRight - currentLeft;
      
      if (newWidth > 0) {
        element.setWidth(newWidth);
        successCount++;
      }
    } catch (e) {
      // Skip elements that don't support these operations
    }
  }
  
  return `✅ Stretched right edge of ${successCount} object(s) to anchor`;
}

/**
 * Stretch Top - Extends objects' top edge to match anchor's top edge
 * 
 * How it works:
 * 1. Identifies the anchor object
 * 2. For each selected object, extends its top edge to align with anchor's top edge
 * 3. Bottom edge stays fixed (object grows/shrinks upward)
 * 
 * Example:
 * - Anchor top edge: 100 points
 * - Object top edge: 150 points, bottom edge: 300 points
 * - After: top edge: 100 points, bottom edge: 300 points (height increased by 50)
 * 
 * Use cases:
 * - Extending objects to a common top boundary
 * - Creating top-aligned blocks with consistent top edge
 * 
 * @returns {string} Status message
 */
function stretchTop() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to stretch';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const targetTop = anchor.getTop();
  
  // Stretch top edge of all non-anchor elements
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      const currentTop = element.getTop();
      const currentBottom = currentTop + element.getHeight();
      const newHeight = currentBottom - targetTop;
      
      if (newHeight > 0) {
        element.setTop(targetTop);
        element.setHeight(newHeight);
        successCount++;
      }
    } catch (e) {
      // Skip elements that don't support these operations
    }
  }
  
  return `✅ Stretched top edge of ${successCount} object(s) to anchor`;
}

/**
 * Stretch Bottom - Extends objects' bottom edge to match anchor's bottom edge
 * 
 * How it works:
 * 1. Identifies the anchor object
 * 2. For each selected object, extends its bottom edge to align with anchor's bottom edge
 * 3. Top edge stays fixed (object grows/shrinks downward)
 * 
 * Example:
 * - Anchor bottom edge: 400 points
 * - Object top edge: 100 points, bottom edge: 300 points
 * - After: top edge: 100 points, bottom edge: 400 points (height increased by 100)
 * 
 * Use cases:
 * - Extending objects to a common bottom boundary
 * - Creating bottom-aligned blocks with consistent bottom edge
 * 
 * @returns {string} Status message
 */
function stretchBottom() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to stretch';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const targetBottom = anchor.getTop() + anchor.getHeight();
  
  // Stretch bottom edge of all non-anchor elements
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      const currentTop = element.getTop();
      const newHeight = targetBottom - currentTop;
      
      if (newHeight > 0) {
        element.setHeight(newHeight);
        successCount++;
      }
    } catch (e) {
      // Skip elements that don't support these operations
    }
  }
  
  return `✅ Stretched bottom edge of ${successCount} object(s) to anchor`;
}

// ============================================================================
// FILL SPACE
// ============================================================================

/**
 * Fill Left - Stretches objects to fill the gap between them and the anchor's right edge
 * 
 * How it works:
 * 1. Identifies the anchor object
 * 2. For each selected object, extends its left edge until it reaches the anchor's right edge
 * 3. Right edge of the object stays fixed
 * 4. If object is already past the anchor, no change occurs
 * 
 * Example:
 * - Anchor right edge: 200 points
 * - Object left edge: 300 points, right edge: 400 points
 * - After: left edge: 200 points, right edge: 400 points (filled 100-point gap)
 * 
 * Use cases:
 * - Filling horizontal space between two objects
 * - Creating connectors between elements
 * - Extending backgrounds to meet boundaries
 * 
 * @returns {string} Status message
 */
function fillLeft() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to fill space';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const anchorRight = anchor.getLeft() + anchor.getWidth();
  
  // Fill left: extend left edge to anchor's right edge
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      const currentLeft = element.getLeft();
      const currentRight = currentLeft + element.getWidth();
      
      // Only fill if there's a gap (object is to the right of anchor)
      if (currentLeft > anchorRight) {
        const newWidth = currentRight - anchorRight;
        element.setLeft(anchorRight);
        element.setWidth(newWidth);
        successCount++;
      }
    } catch (e) {
      // Skip elements that don't support these operations
    }
  }
  
  return successCount > 0 
    ? `✅ Filled left space for ${successCount} object(s) to anchor`
    : '⚠️ No gaps found to fill (objects may already overlap anchor)';
}

/**
 * Fill Right - Stretches objects to fill the gap between them and the anchor's left edge
 * 
 * How it works:
 * 1. Identifies the anchor object
 * 2. For each selected object, extends its right edge until it reaches the anchor's left edge
 * 3. Left edge of the object stays fixed
 * 4. If object is already past the anchor, no change occurs
 * 
 * Example:
 * - Anchor left edge: 300 points
 * - Object left edge: 100 points, right edge: 200 points
 * - After: left edge: 100 points, right edge: 300 points (filled 100-point gap)
 * 
 * Use cases:
 * - Filling horizontal space between two objects
 * - Creating connectors between elements
 * - Extending backgrounds to meet boundaries
 * 
 * @returns {string} Status message
 */
function fillRight() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to fill space';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const anchorLeft = anchor.getLeft();
  
  // Fill right: extend right edge to anchor's left edge
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      const currentLeft = element.getLeft();
      const currentRight = currentLeft + element.getWidth();
      
      // Only fill if there's a gap (object is to the left of anchor)
      if (currentRight < anchorLeft) {
        const newWidth = anchorLeft - currentLeft;
        element.setWidth(newWidth);
        successCount++;
      }
    } catch (e) {
      // Skip elements that don't support these operations
    }
  }
  
  return successCount > 0 
    ? `✅ Filled right space for ${successCount} object(s) to anchor`
    : '⚠️ No gaps found to fill (objects may already overlap anchor)';
}

/**
 * Fill Top - Stretches objects to fill the gap between them and the anchor's bottom edge
 * 
 * How it works:
 * 1. Identifies the anchor object
 * 2. For each selected object, extends its top edge until it reaches the anchor's bottom edge
 * 3. Bottom edge of the object stays fixed
 * 4. If object is already past the anchor, no change occurs
 * 
 * Example:
 * - Anchor bottom edge: 200 points
 * - Object top edge: 300 points, bottom edge: 400 points
 * - After: top edge: 200 points, bottom edge: 400 points (filled 100-point gap)
 * 
 * Use cases:
 * - Filling vertical space between two objects
 * - Creating vertical connectors
 * - Extending backgrounds vertically
 * 
 * @returns {string} Status message
 */
function fillTop() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to fill space';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const anchorBottom = anchor.getTop() + anchor.getHeight();
  
  // Fill top: extend top edge to anchor's bottom edge
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      const currentTop = element.getTop();
      const currentBottom = currentTop + element.getHeight();
      
      // Only fill if there's a gap (object is below anchor)
      if (currentTop > anchorBottom) {
        const newHeight = currentBottom - anchorBottom;
        element.setTop(anchorBottom);
        element.setHeight(newHeight);
        successCount++;
      }
    } catch (e) {
      // Skip elements that don't support these operations
    }
  }
  
  return successCount > 0 
    ? `✅ Filled top space for ${successCount} object(s) to anchor`
    : '⚠️ No gaps found to fill (objects may already overlap anchor)';
}

/**
 * Fill Bottom - Stretches objects to fill the gap between them and the anchor's top edge
 * 
 * How it works:
 * 1. Identifies the anchor object
 * 2. For each selected object, extends its bottom edge until it reaches the anchor's top edge
 * 3. Top edge of the object stays fixed
 * 4. If object is already past the anchor, no change occurs
 * 
 * Example:
 * - Anchor top edge: 300 points
 * - Object top edge: 100 points, bottom edge: 200 points
 * - After: top edge: 100 points, bottom edge: 300 points (filled 100-point gap)
 * 
 * Use cases:
 * - Filling vertical space between two objects
 * - Creating vertical connectors
 * - Extending backgrounds vertically
 * 
 * @returns {string} Status message
 */
function fillBottom() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length < 2) {
    return '⚠️ Please select at least 2 objects to fill space';
  }
  
  // Get anchor element
  const anchor = getAnchorOrFallback(elements);
  const anchorTop = anchor.getTop();
  
  // Fill bottom: extend bottom edge to anchor's top edge
  let successCount = 0;
  for (const element of elements) {
    // Skip the anchor itself
    if (element.getObjectId() === anchor.getObjectId()) {
      continue;
    }
    
    try {
      const currentTop = element.getTop();
      const currentBottom = currentTop + element.getHeight();
      
      // Only fill if there's a gap (object is above anchor)
      if (currentBottom < anchorTop) {
        const newHeight = anchorTop - currentTop;
        element.setHeight(newHeight);
        successCount++;
      }
    } catch (e) {
      // Skip elements that don't support these operations
    }
  }
  
  return successCount > 0 
    ? `✅ Filled bottom space for ${successCount} object(s) to anchor`
    : '⚠️ No gaps found to fill (objects may already overlap anchor)';
}

// ============================================================================
// MAGIC RESIZER
// ============================================================================

/**
 * Show Magic Resizer Dialog - Displays a dialog for user to input resize percentage
 * 
 * This function creates an HTML dialog that prompts the user for a percentage value.
 * The percentage is then used to scale selected objects proportionally.
 * 
 * Examples:
 * - 100% = no change (original size)
 * - 50% = half the current size
 * - 200% = double the current size
 * - 120% = 20% larger
 * 
 * @returns {string} Status message
 */
function showMagicResizerDialog() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: Arial, sans-serif;
          padding: 20px;
          margin: 0;
        }
        h3 {
          margin-top: 0;
          color: #1a73e8;
        }
        .input-group {
          margin: 20px 0;
        }
        label {
          display: block;
          margin-bottom: 8px;
          font-weight: bold;
          color: #333;
        }
        input[type="number"] {
          width: 100%;
          padding: 10px;
          font-size: 16px;
          border: 2px solid #ddd;
          border-radius: 4px;
          box-sizing: border-box;
        }
        input[type="number"]:focus {
          outline: none;
          border-color: #1a73e8;
        }
        .button-group {
          margin-top: 20px;
          display: flex;
          gap: 10px;
        }
        button {
          flex: 1;
          padding: 12px;
          font-size: 14px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-weight: bold;
        }
        .apply-btn {
          background-color: #1a73e8;
          color: white;
        }
        .apply-btn:hover {
          background-color: #1557b0;
        }
        .cancel-btn {
          background-color: #f1f3f4;
          color: #333;
        }
        .cancel-btn:hover {
          background-color: #e0e0e0;
        }
        .examples {
          margin-top: 15px;
          padding: 12px;
          background-color: #f8f9fa;
          border-radius: 4px;
          font-size: 13px;
          color: #555;
        }
        .examples strong {
          color: #333;
        }
        .status {
          margin-top: 15px;
          padding: 10px;
          border-radius: 4px;
          display: none;
        }
        .status.success {
          background-color: #d4edda;
          color: #155724;
          display: block;
        }
        .status.error {
          background-color: #f8d7da;
          color: #721c24;
          display: block;
        }
      </style>
    </head>
    <body>
      <h3>🪄 Magic Resizer</h3>
      
      <div class="input-group">
        <label for="percentage">Resize Percentage:</label>
        <input type="number" id="percentage" value="100" min="1" max="1000" step="1" autofocus>
      </div>
      
      <div class="examples">
        <strong>Examples:</strong><br>
        • 100% = no change (current size)<br>
        • 50% = half size<br>
        • 200% = double size<br>
        • 120% = 20% larger<br>
        • 75% = 25% smaller
      </div>
      
      <div class="button-group">
        <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
        <button class="apply-btn" onclick="applyResize()">Apply</button>
      </div>
      
      <div id="status" class="status"></div>
      
      <script>
        // Allow Enter key to submit
        document.getElementById('percentage').addEventListener('keypress', function(e) {
          if (e.key === 'Enter') {
            applyResize();
          }
        });
        
        function applyResize() {
          const percentage = parseFloat(document.getElementById('percentage').value);
          const statusDiv = document.getElementById('status');
          
          if (isNaN(percentage) || percentage <= 0) {
            statusDiv.className = 'status error';
            statusDiv.textContent = '⚠️ Please enter a valid percentage greater than 0';
            return;
          }
          
          // Disable button while processing
          const applyBtn = document.querySelector('.apply-btn');
          applyBtn.disabled = true;
          applyBtn.textContent = 'Applying...';
          
          // Call server-side function
          google.script.run
            .withSuccessHandler(function(result) {
              statusDiv.className = 'status success';
              statusDiv.textContent = result;
              setTimeout(function() {
                google.script.host.close();
              }, 1500);
            })
            .withFailureHandler(function(error) {
              statusDiv.className = 'status error';
              statusDiv.textContent = '❌ Error: ' + error.message;
              applyBtn.disabled = false;
              applyBtn.textContent = 'Apply';
            })
            .applyMagicResize(percentage);
        }
      </script>
    </body>
    </html>
  `)
    .setWidth(400)
    .setHeight(350);
  
  SlidesApp.getUi().showModalDialog(htmlOutput, 'Magic Resizer');
  return ''; // Dialog handles its own status messages
}

/**
 * Apply Magic Resize - Scales selected objects by the given percentage
 * 
 * How it works:
 * 1. Gets all selected objects
 * 2. For each object, multiplies its width and height by (percentage / 100)
 * 3. Keeps the top-left corner in the same position
 * 
 * Examples:
 * - Object: 200 x 100 points
 * - 50% resize: becomes 100 x 50 points
 * - 200% resize: becomes 400 x 200 points
 * - 120% resize: becomes 240 x 120 points
 * 
 * @param {number} percentage - The percentage to resize by (100 = no change)
 * @returns {string} Status message
 */
function applyMagicResize(percentage) {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const elements = selection ? selection.getPageElementRange()?.getPageElements() : null;
  
  if (!elements || elements.length === 0) {
    return '⚠️ Please select at least 1 object to resize';
  }
  
  const factor = percentage / 100;
  let successCount = 0;
  
  for (const element of elements) {
    try {
      const currentWidth = element.getWidth();
      const currentHeight = element.getHeight();
      
      const newWidth = currentWidth * factor;
      const newHeight = currentHeight * factor;
      
      element.setWidth(newWidth);
      element.setHeight(newHeight);
      successCount++;
    } catch (e) {
      // Skip elements that don't support resize
    }
  }
  
  return `✅ Resized ${successCount} object(s) to ${percentage}% of original size`;
}
