// Global variable for API URL
var BASE_URL = "CHANGE THIS TO YOUR API URL";


// Adds "AI Translator" to the Google Docs menu
function onOpen() {
  DocumentApp.getUi()
    .createMenu("AI Translator")
    .addItem("Open Translator", "openSidebar")
    .addToUi();
}

// Opens the sidebar
function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("AI Translator");
  DocumentApp.getUi().showSidebar(html);
}

function getOptions() {
  try {
    var response = UrlFetchApp.fetch(BASE_URL + "/config", {
      method: "post",
      muteHttpExceptions: true,
      headers: {
        "Content-Type": "application/json"
      }
    });

    Logger.log("Response Code: " + response.getResponseCode());
    Logger.log("Response Text: " + response.getContentText());

    if (response.getResponseCode() !== 200) {
      throw new Error("Error fetching options: " + response.getContentText());
    }

    return JSON.parse(response.getContentText());

  } catch (error) {
    Logger.log("Fetch Error: " + error.message);
    return { languages: [], styles: [], models: [] }; // Return empty data if request fails
  }
}


// Retrieves the selected text in the document
function getSelectedText() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  
  if (!selection) {
    Logger.log("No text selected.");
    return "";
  }

  var text = "";
  var elements = selection.getRangeElements();
  
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i].getElement();
    if (elements[i].isPartial()) {
      text += element.asText().getText().substring(elements[i].getStartOffset(), elements[i].getEndOffsetInclusive() + 1);
    } else {
      text += element.asText().getText();
    }
  }

  Logger.log("Selected Text: " + text);
  return text;
}

// Calls the backend API to translate text
function translateText(text, targetLang, model, style) {
  var payload = {
    text: text,
    target_lang: targetLang,
    model: model,
    style: style
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(BASE_URL + "/translate", options);
    var json = JSON.parse(response.getContentText());
    var translatedText = json.translated_text;

    // Replace selected text with translated text and retain selection
    replaceSelectedText(translatedText);

    return translatedText;
  } catch (e) {
    Logger.log("❌ Error: " + e.toString());
    return "Error contacting translation service.";
  }
}

// Replaces the selected text with translated text and retains the selection
function replaceSelectedText(newText) {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();

  if (!selection) {
    Logger.log("No selection to replace.");
    return;
  }

  var elements = selection.getRangeElements();
  
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i].getElement();
    if (elements[i].isPartial()) {
      var textElement = element.asText();
      var start = elements[i].getStartOffset();
      var end = elements[i].getEndOffsetInclusive();
      
      // Replace only the selected portion of the text
      var originalText = textElement.getText();
      var before = originalText.substring(0, start);
      var after = originalText.substring(end + 1);
      textElement.setText(before + newText + after);
      
      // Retain selection by manually re-selecting the translated text
      var rangeBuilder = doc.newRange();
      rangeBuilder.addElement(textElement, start, start + newText.length - 1);
      doc.setSelection(rangeBuilder.build());
    }
  }
}
