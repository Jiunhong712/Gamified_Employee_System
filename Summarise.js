// Link to google docs: https://docs.google.com/document/d/1sSnpS_MmhmZn9290FmY5CX7R_VGLgSFQJ02s1Ljh_6c/edit?usp=sharing

function onOpen() {
  try {
    var ui = DocumentApp.getUi();
    ui.createMenu('Summarize')
        .addItem('Summarize selected paragraph', 'summarizeSelectedParagraph')
        .addToUi();
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

var API_key = "YOUR_API_KEY_HERE"; 

function generateSummary(paragraph){
  var Model_ID = "gpt-3.5-turbo";
  var maxtokens = 200;
  var temperature = 0.7;
  
  var payload = {
    'model': Model_ID,
    'prompt': 'Please generate a short summary for:\n' +paragraph,
    'temperature': temperature,
    'max_tokens': maxtokens,
    "presence_penalty": 0.5,
    "frequency_penalty": 0.5
  }

  var options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization" : "Bearer " + API_key
    },
    "payload": JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/completions", options);
  var summary = JSON.parse(response.getContentText());
  return summary.choices[0].text.trim();
}

function summarizeSelectedParagraph() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) {
    DocumentApp.getUi().alert('No text selected');
    return;
  }
  
  var text = selection.getRangeElements().map(function(rangeElement) {
    return rangeElement.getElement().asText().getText();
  }).join('\n');
  
  var summary = generateSummary(text);
  
  DocumentApp.getActiveDocument().getBody().appendParagraph("Summary");
  DocumentApp.getActiveDocument().getBody().appendParagraph(summary);
}

