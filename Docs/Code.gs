URL = "https://9171-171-239-138-199.ngrok-free.app"

// Add a menu item to Google Docs when the document opens
function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Translation Add-on')
    .addItem('Show Sidebar', 'showSidebar')
    .addToUi();
}

// Display the sidebar with the translation UI
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Translation Add-on');
  DocumentApp.getUi().showSidebar(html);
}

// Function to call the FastAPI /chat endpoint
function callTranslationAPI(text, lang_src, lang_des, style) {
  const url = URL + '/chat'; // Replace with your actual FastAPI URL
  const payload = {
    lang_src: lang_src,
    lang_des: lang_des,
    text: text,
    style: style
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    return json.message;
  } catch (e) {
    return 'Error: Unable to connect to the translation service.';
  }
}

// Translate the currently selected text
function translateSelectedText(lang_src, lang_des, style) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) {
    return 'No text selected.';
  }
  let text = '';
  const rangeElements = selection.getRangeElements();
  for (let i = 0; i < rangeElements.length; i++) {
    const element = rangeElements[i].getElement();
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      text += element.asParagraph().getText() + '\n';
    } else if (element.getType() === DocumentApp.ElementType.TEXT) {
      const textElement = element.asText();
      if (rangeElements[i].isPartial()) {
        const start = rangeElements[i].getStartOffset();
        const end = rangeElements[i].getEndOffsetInclusive() + 1;
        text += textElement.getText().substring(start, end);
      } else {
        text += textElement.getText();
      }
    }
  }
  if (text.trim() === '') {
    return 'No text selected.';
  }
  return callTranslationAPI(text, lang_src, lang_des, style);
}

// Replace the selected text with the translated text
function replaceSelectedText(translatedText) {
  // Get the active Google Docs document
  const doc = DocumentApp.getActiveDocument();
  
  // Get the current selection
  const selection = doc.getSelection();
  
  // Check if there’s no selection
  if (!selection) {
    DocumentApp.getUi().alert('No text selected. Please select some text and try again.');
    return;
  }
  
  // Get the range elements (parts of the document that are selected)
  const rangeElements = selection.getRangeElements();
  if (rangeElements.length === 0) {
    DocumentApp.getUi().alert('No elements found in the selection.');
    return;
  }
  
  // For now, support only selections within a single paragraph
  if (rangeElements.length > 1 || rangeElements[0].getElement().getType() !== DocumentApp.ElementType.TEXT) {
    DocumentApp.getUi().alert('Please select text within a single paragraph.');
    return;
  }
  
  // Get the selected text element and its range
  const rangeElement = rangeElements[0];
  const textElement = rangeElement.getElement().asText();
  const start = rangeElement.getStartOffset();
  const end = rangeElement.getEndOffsetInclusive();
  
  // Replace the selected text with the translated text
  textElement.deleteText(start, end);
  textElement.insertText(start, translatedText);
}

// Translate and replace the entire document's text
function translateAndReplaceWholeDocument(lang_src, lang_des, style) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const text = body.getText();
  if (!text.trim()) {
    DocumentApp.getUi().alert('Document is empty.');
    return;
  }
  const translatedText = callTranslationAPI(text, lang_src, lang_des, style);
  body.setText(translatedText);
}