// Script to create section word counts in document headings
// Author: Ben Coleman (ben[dot]coleman[at]alum[dot]utoronto[dot]ca)
// Github repository: [add link]


// PREFERENCES
// text highlight colors
const underColor = '#b6d7a8'; // under the word count (default: pastel green)
const closeColor = '#ffe599'; // close to the word count (default: pastel yellow)
const overColor = '#ea9999'; // over the word count (default: pastel red)

// regular expressions
// default works with "(<wordcount>/<wordlimit>w)" or "(<wordcount>/<wordlimit>words)"
const rxWordLimit = /(\(\s*\d*\/*)(\d+)(?:w|words)\)/; // entire heading wordcount section
const rxBeginning = rxWordLimit, rxEnd = /\/(\d+)(?:w|words)\)/; // used for text highlighting
const rxSlashAbsent = /\(\d+(?:w|words)\)/; // used to handle cases without a slash


function headingWordCount() {
  currentNode = DocumentApp.getActiveDocument().getBody().getChild(0);
  var runningCount = 0;
  var previousHeading;

// special case for first node in document
  if (currentNode.getHeading() != 'NORMAL') {
    previousHeading = currentNode;
  } else if (currentNode.getHeading() == 'NORMAL') {
    runningCount += countWords(currentNode);
  }

// main processing loop
  do {
    currentNode = currentNode.getNextSibling();
    // headings
    if (currentNode.getHeading() != 'NORMAL') {
      headingUpdate(previousHeading, runningCount);
      runningCount = 0;
      previousHeading = currentNode;
    // normal paragraphs
    } else if (currentNode.getHeading() == 'NORMAL') {
      runningCount += countWords(currentNode);
    }
    // horizontal rules
    if (currentNode.findElement(DocumentApp.ElementType.HORIZONTAL_RULE) != null) {
      runningCount = 0;
    }
  } while (currentNode.isAtDocumentEnd() == false);
  headingUpdate(previousHeading, runningCount);
}


// Count words - from Jed Grant (source: https://stackoverflow.com/questions/33338667/function-for-word-count-in-google-docs-apps-script)
function countWords(p) {
  var s = p.getText();

  //check for empty nodes
  if (s.length === 0)
      return 0;
  //A simple \n replacement didn't work, neither did \s not sure why
  s = s.replace(/\r\n|\r|\n/g, " ");
  //In cases where you have "...last word.First word..."
  //it doesn't count the two words around the period.
  //so I replace all punctuation with a space
  var punctuationless = s.replace(/[.,\/#!$%\^&\*;:{}=\_`~()"?“”…]/g," ");
  //Finally, trim it down to single spaces (not sure this even matters)
  var finalString = punctuationless.replace(/\s{2,}/g," ");
  //Actually count it
  var count = finalString.trim().split(/\s+/).length;
  return count;
}


// Update and style headings
function headingUpdate(heading, wordCount) {
  // extract word limit
  wordLimit = heading.getText().match(rxWordLimit);
  slashAbsent = heading.getText().match(rxSlashAbsent);

  // write in word count with appropriate formatting
  if (wordLimit != null) {
    // add new word count to heading
    if (slashAbsent != null) {
      var wcInsert = heading.getText().search(slashAbsent);
      heading.editAsText().insertText(wcInsert,wordCount + "/");
      // TODO temporary fix - index of wordLimit is 1, not 2 when slash absent
      wordLimit = heading.getText().match(rxWordLimit);
    } else {
      heading.setText(heading.getText().replace(wordLimit[1], "(" + wordCount + "/"));
    }

    // determine string position of updated wordcount
    var text = heading.getText();
    var srBeginning = text.search(rxBeginning);
    var srEnd = text.search(rxEnd);

    // build format objects
    var hlUnder = {}, hlClose = {}, hlOver = {};
    hlUnder[DocumentApp.Attribute.BACKGROUND_COLOR] = underColor;
    hlClose[DocumentApp.Attribute.BACKGROUND_COLOR] = closeColor;
    hlOver[DocumentApp.Attribute.BACKGROUND_COLOR] = overColor;

    // calculate whether word count is under, close to, or over the limit
    var threshold = Math.max(Math.round(0.9 * wordLimit[2]),wordLimit[2] - 20);
    if (wordCount < threshold) {
      highlightStyle = hlUnder;
    } else if (wordCount >= threshold && wordCount <= wordLimit[2]) {
      highlightStyle = hlClose;
    } else if (wordCount > wordLimit[2]) {
      highlightStyle = hlOver;
    }

    // apply highlight formatting
    heading.editAsText().setAttributes(srBeginning + 1, srEnd - 1, highlightStyle);
  }
  // return
}


// MENU & TRIGGER GENERATION
// Check if document length has changed before running time trigger
function isEdited() {
  var myDoc = DocumentApp.getActiveDocument().getBody();
  var docText = myDoc.editAsText().getText();
  var docLen= docText.length;
  if(docLen!= PropertiesService.getDocumentProperties().getProperty('docLen'))
  {
    headingWordCount();
    PropertiesService.getDocumentProperties().setProperty('docLen', docLen)
  }
}

// Create custom menu to run utility
function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Word Count')
      .addItem('Update Word Counts', 'headingWordCount')
      .addToUi();
  createTimeTrigger();
}

// Create time-based trigger
function createTimeTrigger() {
  // Trigger every minute
  ScriptApp.newTrigger('isEdited')
    .timeBased()
    .everyMinutes(1)
    .create();
}
