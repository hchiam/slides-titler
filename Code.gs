/**
 * @OnlyCurrentDoc Adds progress bars to a presentation.
 */

function onOpen() {
  SlidesApp.getUi()
    .createMenu("Slides-titler")
    .addItem("Add title to all slides", "createTitles")
    .addItem("Remove title from all slides", "deleteTitles")
    .addToUi();
}

/**
 * Create a title on every slide.
 */
function createTitles() {
  const titleText = getTitleText();
  if (!titleText) return;
  deleteTitles(); // Delete any existing progress bars
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    const x = 0;
    const y = 0;
    let width = 20 + titleText.length * 5;
    if (width > presentation.getPageWidth()) {
      width = presentation.getPageWidth();
    }
    const isASlide = width > 0;
    if (isASlide) {
      const TITLE_HEIGHT = 15; // px
      const TITLE_FONT_SIZE = 7;
      const title = slides[i]
        .insertTextBox(titleText || "(Title)", x, y, width, TITLE_HEIGHT)
        .setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      title.getBorder().setTransparent();
      title
        .getText()
        .getTextStyle()
        .setForegroundColor("#ffffff")
        .setFontSize(TITLE_FONT_SIZE)
        .setBold(true);
      const backgroundColor = "#000000";
      const alpha = 0.5;
      const TITLE_ID = "TITLE_ID";
      title.getFill().setSolidFill(backgroundColor, alpha);
      title.setLinkUrl(TITLE_ID);
    }
  }
}

function getTitleText() {
  const ui = SlidesApp.getUi();
  const result = ui.prompt(
    "What's the title to put on every slide?",
    "Please enter your title here:",
    ui.ButtonSet.OK_CANCEL
  );
  const button = result.getSelectedButton();
  const titleText = result.getResponseText();
  const buttonType = button == ui.Button.OK ? "OK" : "Cancel";
  const hitOk = buttonType === "OK";
  if (hitOk) {
    if (titleText) {
      return titleText;
    } else {
      ui.alert("You didn't enter text for the title. Cancelling.");
      return "";
    }
  }
  return "";
}

/**
 * Deletes all the titles that were auto-generated (finds by link URL matching an ID).
 */
function deleteTitles() {
  const TITLE_ID = "TITLE_ID";
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  for (let i = 0; i < slides.length; ++i) {
    const elements = slides[i].getPageElements();
    for (let j = 0; j < elements.length; ++j) {
      const el = elements[j];
      if (
        el.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
        el.asShape().getLink() &&
        el.asShape().getLink().getUrl() === TITLE_ID
      ) {
        el.remove();
      }
    }
  }
}
