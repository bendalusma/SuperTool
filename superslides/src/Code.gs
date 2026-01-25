/**
 * onOpen()
 *
 * Runs automatically when a Slides presentation opens.
 * Adds a custom menu so you can open your sidebar.
 */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('SuperSlides')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

/**
 * showSidebar()
 *
 * Creates and displays the sidebar UI.
 * The sidebar HTML lives in Sidebar.html.
 */
function showSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('Sidebar')
    .setTitle('SuperSlides');

  SlidesApp.getUi().showSidebar(html);
}

/**
 * setAnchor()
 *
 * MVP behavior:
 * - Reads the current user selection in Slides.
 * - Expects at least one selected page element (shape, image, etc.).
 * - Stores the selected element's object ID in DocumentProperties.
 *
 * Notes:
 * - DocumentProperties are stored per presentation (document), not per user.
 *   For MVP that's fine.
 * - Later we can switch to UserProperties if you want each user to keep their own anchor.
 */
function setAnchor() {
  const selection = SlidesApp.getActivePresentation().getSelection();

  // If nothing is selected, getPageElementRange() can be null.
  const range = selection.getPageElementRange();
  if (!range) {
    return 'No selection — please select a shape first.';
  }

  const elements = range.getPageElements();
  if (!elements || elements.length === 0) {
    return 'No elements selected — please select a shape first.';
  }

  // For setting the anchor, we intentionally use the FIRST selected element.
  // This makes the anchor deterministic: user selects one object then clicks Set Anchor.
  const anchorId = elements[0].getObjectId();

  // Persist anchorId so other functions can retrieve it later.
  PropertiesService.getDocumentProperties().setProperty('anchorId', anchorId);

  return 'Anchor set.';
}

/**
 * clearAnchor()
 *
 * Removes the stored anchor from DocumentProperties.
 */
function clearAnchor() {
  PropertiesService.getDocumentProperties().deleteProperty('anchorId');
  return 'Anchor cleared.';
}

/**
 * getAnchorStatus()
 *
 * Called by the sidebar on load.
 * Returns a human-readable string we display in the sidebar.
 */
function getAnchorStatus() {
  const anchorId = PropertiesService.getDocumentProperties().getProperty('anchorId');
  return anchorId ? 'Anchor is set.' : 'No anchor set.';
}

/**
 * getAnchorOrFallback(elements)
 *
 * Resolves which element should be used as the anchor.
 *
 * Rules:
 * 1) If an anchorId exists AND is part of the current selection → use it
 * 2) Otherwise, fall back to the last element in the selection array
 *
 * @param {PageElement[]} elements
 * @return {PageElement}
 */
function getAnchorOrFallback(elements) {
    const props = PropertiesService.getDocumentProperties();
    const anchorId = props.getProperty('anchorId');
  
    if (anchorId) {
      const match = elements.find(el => el.getObjectId() === anchorId);
      if (match) {
        return match;
      }
    }
  
    // Fallback behavior (intentionally imperfect for MVP testing)
    return elements[elements.length - 1];
  }

  /**
 * alignLeft()
 *
 * Aligns all selected elements to the left edge of the anchor element.
 *
 * IMPORTANT:
 * - Always returns a string so the sidebar can update its status reliably.
 * - Uses try/catch per element so one problematic element doesn't hang the whole operation.
 */
function alignLeft() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const range = selection.getPageElementRange();

  if (!range) {
    return 'No page elements selected. Click shapes on the slide canvas first.';
  }

  const elements = range.getPageElements();

  // Need at least 2 elements to align
  if (!elements || elements.length < 2) {
    return 'Select at least 2 shapes to align.';
  }

  // Resolve anchor using our MVP rules
  const anchor = getAnchorOrFallback(elements);
  const anchorLeft = anchor.getLeft();

  let moved = 0;
  let failed = 0;

  // Align all OTHER elements to anchor's left edge
  elements.forEach(el => {
    if (el.getObjectId() === anchor.getObjectId()) {
      return; // skip anchor itself
    }

    try {
      el.setLeft(anchorLeft);
      moved += 1;
    } catch (e) {
      // If a specific element type cannot be moved (rare), don't break everything.
      failed += 1;
    }
  });

  if (failed > 0) {
    return `Aligned ${moved} element(s). ${failed} element(s) could not be moved.`;
  }

  return `Aligned ${moved} element(s) to the anchor’s left edge.`;
}