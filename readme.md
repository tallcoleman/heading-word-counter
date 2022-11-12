# Heading Word Counter

A Google Docs script for grant writers that gives a word count for the text beneath each heading if the heading indicates a word limit.

![](https://github.com/tallcoleman/heading-word-counter/blob/main/assets/example.png)

## Installation

To use the script, you will have to add it to your Google Doc. The steps are as follows:

1. Open the script editor from your Google Doc (Tools > Script editor).

2. Copy the [headingwordcounter.gs script](https://github.com/tallcoleman/heading-word-counter/blob/main/headingwordcounter.gs), paste it into the script editor, and select 'Save'.

3. At this point, it is also a good idea to rename the script project from 'Untitled project' to a name that is the same as or similar to your Google Doc. You can rename the project by selecting its name at the top of your screen.

4. To the right of the icon is a function drop-down (it will probably say 'headingWordCount'). Open this drop-down, select 'createOnOpenTrigger', and then select 'Run'. This starts the automatic count which updates every minute, and ensures the automatic count re-starts if needed when the document is opened.

5. Since this is the first time you are using the script, it will ask you for permissions. (If you are unfamiliar with Google Apps Script permissions, you can [read a short explainer below](#google-apps-script-permissions-why-does-google-say-this-script-is-unsafe).)

6. After selecting "Continue", you may come to a screen that says "Google hasn't verified this app". If you are presented with this screen, select "Advanced" at the bottom-left and then "Go to Untitled project (unsafe)".

7. When you reach the permissions screen, select "Allow".

8. You should see a new menu named "Word Count" in your Google Docs window.

9.  The script should now be fully installed. If you want to run it right away, select Word Count > Update Word/Character Counts.


## Using the Script

The script counts all the words or characters below a heading until it reaches the next heading. In order to create a heading, you have to use [document styles](https://support.google.com/docs/answer/116338) (e.g. Title, Heading 1, Heading 2, etc.)

The script will only create and update a word count if you put a word or character limit in brackets somewhere in the heading text. For example, for a section with a word limit of 250 words, you could add either `(250w)` or `(250words)` anywhere in your heading. The format for character counts would be similar: `(1000c)`, `(1000char)`, and `(1000chars)` will all work by default.

After it has run, the script will add or update the relevant count in these brackets. For example, if you have written 150 words, the text in parentheses will now say `(150/250w)`. The script will also highlight the count with a color to indicate if you are below, close to, or above the word limit.

### Running the Script

The script will update the word counts in your document:
* Automatically every minute; or
* Immediately after selecting Word Count > Update Word/Character Counts in the menu bar.

### Automatic Update Timeout

If the document has not been edited in a while (30 minutes by default) the script will stop automatically updating word counts. The timeout can be restarted by reloading the document or by selecting Word Count > Update Word/Character Counts in the menu bar.

This timeout feature helps to minimize the amount that the script uses up your [daily quotas for Apps Script services](https://developers.google.com/apps-script/guides/services/quotas?hl=en).

Due to the limitations of the Apps Script service, the script is not currently able to automatically update the word counts more than once a minute.

### Excluding Text From the Word Count

If you need to put instructions or notes after your heading but do not want them included in the word count, you can put a horizontal line just after the instructions (Insert > Horizontal line).

The horizontal line will reset the word count to zero, so note that if you have multiple horizontal lines in your section, the word count will only include the text after the _last_ horizontal line before the next heading.

### Customization
The script is highly customizable using the variables in the preferences section. Options include:

* Providing your own highlight colours
* Setting a custom threshold for getting "close" to the word or character limit
* Creating your own syntax for including the word count in headings
* Changing the automatic update timeout

The only caveat is that the values in the text matching section need to be regular expression–friendly. For example, if you wanted to use `[250w]` instead of `(250w)`, you would have to use the following, since `[` and `]` are special characters:

```javascript
const textEncapsulators = [raw`\[`, raw`\]`];
```

(Note: `raw` is used in the script as a concise alias for [`String.raw`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/raw). If you want to use regular strings, you can double escape special characters instead (e.g. `'\\['` or `'\\]'`).)

## Uninstalling the Script

The easiest way to completely remove the script, including its triggers and permissions, is to go to the "Overview" menu in the script editor (`(i)` icon), and then select "Delete Project Forever" at the top-right of the page (trash can icon).

If you are using other scripts in your document, you can simply delete the script code and remove its triggers in the "Triggers" menu of the script editor (clock icon). You should be able to determine the correct triggers to remove based on the function column:

* `onOpenActions` for the trigger that runs when the document is opened or reloaded
* `runCount` for the trigger that runs every minute (if it hasn't already been removed by the timeout)


## Google Apps Script Permissions (Why does Google say this script is unsafe?)
Google Apps use the same security system for scripts that you write (or copy in) yourself and add-ons created by third parties. The permissions system is therefore designed to be very cautious so that users do not give third parties unintended permissions for their Google Account.

The permissions used by the script are required for the following reasons:

**See, create, and edit all Google Docs documents you have access to**

This permission is required by the functions that work with your document (e.g. counting words, updating headings with word counts). The script does not read or use any document other than the one you add it to.

**Allow this application to run when you are not present**

This permission is required by a trigger that runs every minute to update the word counts in your document while you are writing. This trigger does not do anything while the document is closed, since it only acts if the number of words in the document has changed since it last ran. It is automatically removed after each timeout period.


## Future Plans

* Improve the word counting algorithm and add customization options to account for grant portals that use slightly different word counting methods.

* Ensure that the script correctly handles all of the content types allowable in Google Docs (e.g. tables, formulas); add and test error handling.

* Improve triggers if possible (see if there is a way to update more frequently, simplify process for multiple documents).

* Explore option to publish as an add-on.
