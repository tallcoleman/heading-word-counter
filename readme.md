# Heading Word Counter

A Google Docs script for grant writers that gives a word count for the text beneath each heading if the heading indicates a word limit.

![](https://github.com/tallcoleman/heading-word-counter/blob/main/assets/example.png)

## Installation

To use the script, you will have to add it to your Google Doc. The steps are as follows:

1. Open the script editor from your Google Doc (Tools > Script editor).

2. Copy the [headingwordcounter.gs script](https://github.com/tallcoleman/heading-word-counter/blob/main/headingwordcounter.gs), paste it into the script editor, and select 'Save'.

3. At this point, it is also a good idea to rename the script project from 'Untitled project' to a name that is the same as or similar to your Google Doc. You can rename the project by selecting its name at the top of your screen.

4. To the right of the icon is a function drop-down (it will probably say 'headingWordCount'). Open this drop-down, select 'createTimeTrigger', and then select 'Run'. This will create a trigger to run the script every minute.

5. Since this is the first time you are using the script, it will ask you for permissions. (If you are unfamiliar with Google Apps Script permissions, you can [read a short explainer below](#google-apps-script-permisions-why-does-google-say-this-script-is-unsafe).)

6. After selecting "Continue", you may come to a screen that says "Google hasn't verified this app". If you are presented with this screen, select "Advanced" at the bottom-left and then "Go to Untitled project (unsafe)".

7. When you reach the permissions screen, select "Allow".

8. Use the reload button in your browser to reload your Google Docs page. After the document has fully loaded again, you should see a new menu named "Word Count".

9. The script should now be fully installed. If you want to run it right away, select Word Count > Update Word Counts.


## Using the Script

The script counts all the words below a heading until it reaches the next heading. In order to create a heading, you have to use [document styles](https://support.google.com/docs/answer/116338) (e.g. Title, Heading 1, Heading 2, etc.)

The script will only create and update a word count if you put a word limit in brackets somewhere in the heading text. For example, for a section with a word limit of 250 words, you could add either `(250w)` or `(250words)` anywhere in your heading.

After it has run, the script will add or update a word count in these brackts. For example, if you have written 150 words, the text in parentheses will now say `(150/250w)`. The script will also highlight the word count with a color to indicate if you are below, close to, or above the word limit.

The script will update the word counts in your document:
* Every minute; or
* Immediately after selecting Word Count > Update Word Counts in the menu bar.

Due to the limitations of the Apps Script service, the script is not currently able to automatically update the word counts more than once a minute.

### Excluding Instruction Text From the Word Count

If you need to put instructions after your heading but do not want them included in the word count, you can put a horizontal line just after the instructions (Insert > Horizontal line).

The horizontal line will reset the word count to zero, so note that if you have multiple horizontal lines in your section, the word count will only include the text after the _last_ horizontal line before the next heading.

### Customization
You can customize the highlight colors used by changing the hexadecimal color codes at the top of the script.


## Uninstalling the Script

To remove the script:

1. Delete the script using the script editor

2. Using the menu on the left of the script editor, go to the Triggers pane (alarm clock icon). Delete the trigger listed on the pane using the three-dot menu on its right.

3. Go to [myaccount.google.com/permissions](https://myaccount.google.com/permissions) and remove the project from your permissions list. The project will have the name you gave it in step 3 of the installation. If you did not rename the project, it will be named 'Untitled project'. Select "REMOVE ACCESS" to remove the permissions for the script.


## Google Apps Script Permisions (Why does Google say this script is unsafe?)
Google Apps use the same security system for scripts that you write (or copy in) yourself and add-ons created by third parties. The permissions system is therefore designed to be very cautious so that users do not give third parties unintended permissions for their Google Account.

The permissions used by the script are required for the following reasons:

**See, create, and edit all Google Docs documents you have access to**

This permission is required by the functions that work with your document (e.g. counting words, updating headings with word counts). The script does not read or use any document other than the one you add it to.

**Allow this application to run when you are not present**

This permission is required by a trigger that runs every minute to update your word counts while you are writing. At the moment, I have not added the work-around required to stop the trigger when the document is not open. However, the trigger does nothing while the document is closed. (The script checks to see if the document length has changed before running).

## Future Plans

* Improve the word counting algorithm and add customizability to account for grant portals that use different word counting methods than Google Docs.

* Check that the script correctly handles all of the content types allowable in Google Docs (e.g. tables, formulas); add and test error handling.

* Improve triggers (stop time trigger when document is closed; explore using [`Files: watch`](https://developers.google.com/drive/api/v3/reference/files/watch) to approximate onEdit trigger available in Google Sheets).

* Allow customizable syntax for word limit, improve text handling, and add options for character counts in addition to word counts.

* Explore option to publish as an add-on.
