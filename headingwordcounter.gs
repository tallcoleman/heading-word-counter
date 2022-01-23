// Script to create section word/character counts in document headings
// Author: Ben Coleman (ben[dot]coleman[at]alum[dot]utoronto[dot]ca)
// Github repository: https://github.com/tallcoleman/heading-word-counter

// prefix for raw strings
const raw = String.raw;

// PREFERENCES
// text highlight colors in hex format
// under the word/character count (default: pastel green)
const underColor = '#b6d7a8';
// close to the word/character count (default: pastel yellow)
const closeColor = '#ffe599';
// over the word/character count (default: pastel red) 
const overColor = '#ea9999'; 

// threshold for close colour ('exact' is used for longer sections)
// 'exact' is multiplied by average word length for character counts
const thresholdUnits = {
    proportion: .1,          // % of words left
    exact: 20,               // number of words left
}
const averageWordLength = 5;

// text matching
// needs to be in a format that works in a regular expression
const textEncapsulators = [raw`\(`, raw`\)`];         // default '(' and ')'
const textSeparators = [raw`\/`];                     // default '/'
const textWordLimitTokens = ['w', 'words'];
const textCharLimitTokens = ['c', 'char', 'chars'];
const textMaxLimitTokens = ['max', 'maximum'];


// INITIALIZATION
// create regex string
let reComponents = [
    `(?<unitCountPrefix>` +
        textEncapsulators[0],
    `)` +
    raw`(?<unitCount>\d+)??` +
    `(?<unitCountSuffix>`,
    `(?<trailingSeparator>${textSeparators.join('|')})?`,
        `(?:` +
            raw`(?<wordLimit>\d+)`,
            `(?:${textWordLimitTokens.join('|')})` +
            `|` +
            raw`(?<charLimit>\d+)`,
            `(?:${textCharLimitTokens.join('|')})` +
        `)`,
        `(?:` +
            `(?:${textSeparators.join('|')})`,
            raw`(?<maxLimit>\d+)`,
            `(?:${textMaxLimitTokens.join('|')})` +
        `)?`,
        textEncapsulators[1] +
    `)`

];

const reIndicator = new RegExp(reComponents.join(raw`\s*?`));


// MAIN FUNCTIONS
// Iterate through paragraphs
function parProcess(currentNode, runningCount = {words: 0, chars: 0}, previousHeading) {
    // skips tables
    if (currentNode.getHeading()) {
        // headings
        if (currentNode.getHeading() != 'NORMAL') {
            if (previousHeading) headingUpdate(previousHeading, runningCount);
            runningCount = {words: 0, chars: 0};
            previousHeading = currentNode;
        // normal paragraphs
        } else if (currentNode.getHeading() == 'NORMAL') {
            runningCount.words += countWords(currentNode);
            runningCount.chars += currentNode.getText().length;
        }

        // reset unit count at horizontal rules
        if (currentNode.findElement(DocumentApp.ElementType.HORIZONTAL_RULE)) {
            runningCount = {words: 0, chars: 0};
        }
    }

    // stop if at end of doc
    if (currentNode.isAtDocumentEnd()) {
        headingUpdate(previousHeading, runningCount);
    } else {
        return parProcess(currentNode.getNextSibling(), runningCount, previousHeading);
    }

    return;
}


// Count words - from Jed Grant (source: https://stackoverflow.com/questions/33338667/function-for-word-count-in-google-docs-apps-script)
function countWords(p) {
  var s = p.getText();

  // check for empty nodes
  if (s.length === 0)
      return 0;
  // A simple \n replacement didn't work, neither did \s not sure why
  s = s.replace(/\r\n|\r|\n/g, " ");
  // In cases where you have "...last word.First word..."
  // it doesn't count the two words around the period.
  // so I replace all punctuation with a space
  var punctuationless = s.replace(/[.,\/#!$%\^&\*;:{}=\_`~()"?“”…]/g," ");
  // Finally, trim it down to single spaces (not sure this even matters)
  var finalString = punctuationless.replace(/\s{2,}/g," ");
  // Actually count it
  var count = finalString.trim().split(/\s+/).length;
  return count;
}


// Update and style headings
function headingUpdate(heading, counts) {
    // extract indicator values
    let matches = reIndicator.exec(heading.getText());
    if (!matches) return;

    // indicator is only valid if it specifies a unit maximum
    let unitType, unitLimit, unitCount;
    if (matches.groups.wordLimit) {
        unitType = 'words';
        unitLimit = matches.groups.wordLimit;
    } else if (matches.groups.charLimit) {
        unitType = 'chars';
        unitLimit = matches.groups.charLimit;
    } else {
        return;
    }
    unitCount = counts[unitType]

    // write in unit count with appropriate formatting
    // add new unit count to heading
    let trailingSeparator = !matches.groups.trailingSeparator ? textSeparators[0].replace('\\','') : '';
    let newHeadingText = heading.getText().replace(reIndicator, `$<unitCountPrefix>${unitCount}${trailingSeparator}$<unitCountSuffix>`);
    heading.setText(newHeadingText);

    // determine string position of updated unit count
    let unitCountBegins = matches.index + matches.groups.unitCountPrefix.length;
    let unitCountEnds = unitCountBegins + unitCount.toString().length;

    // build format objects
    let hlUnder = {}, hlClose = {}, hlOver = {};
    hlUnder[DocumentApp.Attribute.BACKGROUND_COLOR] = underColor;
    hlClose[DocumentApp.Attribute.BACKGROUND_COLOR] = closeColor;
    hlOver[DocumentApp.Attribute.BACKGROUND_COLOR] = overColor;

    // calculate whether unit count is under, close to, or over the limit
    let threshold = Math.max(
        Math.round((1 - thresholdUnits.proportion) * unitLimit),
        unitLimit - thresholdUnits.exact * (unitType === 'chars' ? averageWordLength : 1)
        );
    if (unitCount < threshold) {
    highlightStyle = hlUnder;
    } else if (unitCount >= threshold && unitCount <= unitLimit) {
    highlightStyle = hlClose;
    } else if (unitCount > unitLimit) {
    highlightStyle = hlOver;
    }

    // apply highlight formatting
    heading.editAsText().setAttributes(unitCountBegins, unitCountEnds - 1, highlightStyle);

    return;
}


// MENU & TRIGGER GENERATION
// Check if document length has changed before running time trigger
function isEdited() {
  let myDoc = DocumentApp.getActiveDocument().getBody();
  let docText = myDoc.editAsText().getText();
  let docLen = docText.length;
  if(docLen != PropertiesService.getDocumentProperties().getProperty('docLen'))
  {
    parProcess(myDoc.getChild(0));
    PropertiesService.getDocumentProperties().setProperty('docLen', docLen)
  }
}

// Create custom menu to run utility
function onOpen() {
  let ui = DocumentApp.getUi();
  ui.createMenu('Word Count')
      .addItem('Update Word Counts', 'isEdited')
      .addToUi();
}

// Create time-based trigger
function createTimeTrigger() {
  // Trigger every minute
  ScriptApp.newTrigger('isEdited')
    .timeBased()
    .everyMinutes(1)
    .create();
}