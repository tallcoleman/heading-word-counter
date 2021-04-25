# Heading Word Counter

A Google Docs script for grant writers that gives a word count for the text beneath each heading in a document as long as the heading indicates a word limit.

![](/blob/master/example.png)

## Installation

To use the script, you will have to add the script to your Google Doc. The steps are as follows:

1. Open the script editor (Tools > Script editor)

2. Copy the [script from this repository named headingwordcounter.gs](/blob/master/headingwordcounter.gs). Paste it into the script editor and then save the script using the menu at the top of the script editor. Once the script is saved, you can close the script editor tab.

3. Use the reload button in your browser to reload the Google Docs page. After the document has fully loaded, you should see a new menu named "Word Count".

4. Select Word Count > Update Word Counts to run the script for the first time. (If you are unfamiliar with Google Apps Script permissions, you can [read a short explainer below](#why-does-google-say-this-script-is-unsafe).) You will have to authorize the script. After selecting "Continue", you will come to a screen that says "Google hasn't verified this app" (you may have to select your google account first).

5. Select "Advanced" at the bottom-left and then "Go to Untitled project (unsafe)". On the following permissions screen, select "Allow".

6. The script should now be running, and will run when you use your document in the future. If you want to run it right away, select Word Count > Update Word Counts again.

## Using the Script

The script counts all the words below a heading until it reaches the next heading. In order to create a heading, you have to use [document styles](https://support.google.com/docs/answer/116338) (e.g. Title, Heading 1, Heading 2, etc.)

The script will only create and update a word count if you put a word limit in brackets somewhere in the heading text. For example, for a section with a word limit of 250 words, you could add either `(250w)` or `(250words)` anywhere in your heading.

After it has run the first time, the script will add a word count, e.g. if you have written 150 words, the text in parentheses will now say `(150/250w)`. The script will also highlight the word count with a color to indicate if you are below, close to, or above the word limit.

The script will update:
* Every minute after you have added or deleted words (i.e. if the length of the document changes)
* Immediately after selecting Word Count > Update Word Counts in the menu bar

Due to the limitations of the Apps Script API, it is not possible for the script to automatically run more than once a minute.

### Customization
You can customize the highlight colors used by changing the hexadecimal color codes at the top of the script.

![](/blob/master/exampleclip.m4v)


## Uninstalling the Script

To remove the script:

1. Delete the script from the script editor

2. Using the menu on the left of the script editor, go to the Triggers pane (the icon looks like an alarm clock). Delete the trigger listed on the pane using the three-dot menu on its right.

3. Go to [myaccount.google.com/permissions](https://myaccount.google.com/permissions) and remove the script from your permissions list. It will likely be named "Untitled project" unless you changed this project name in the script editor. Select "REMOVE ACCESS" to remove the permissions for the script.

(If you have multiple items in your permissions list named "Untitled project", you should be able to figure out which is the right one by looking at "Access given on:").


## Google Apps Script Permisions (Why does Google say this script is unsafe?)
Google Apps use the same security system for scripts that you write (or copy in) yourself and add-ons created by third parties. The permissions system is therefore designed to be very cautious so that users do not give third parties unintended permissions for their Google Account.

The permissions used by the script are required for the following reasons:

**See, create, and edit all Google Docs documents you have access to**

This permission is required by the functions that work with your document (e.g. counting words, updating headings with word counts). The script does not read or use any document other than the one you add it to.

**Allow this application to run when you are not present**

This permission is required by a trigger that runs every minute to update your word counts while you are writing. At the moment, I have not added the work-around required to stop the trigger when the document is not open. However, the trigger does nothing while the document is closed since the document length is not changing.

## Future Plans

* Improve the word counting algorithm and add customizability to account for grant portals that use different word counting methods than Google Docs.

* Check that the script correctly handles all of the content types allowable in Google Docs (e.g. tables, formulas); add and test error handling.

* Add a feature to reset the running count after Horizontal Line objects to allow for instruction text to come after a heading but not be counted.

* Improve triggers (stop time trigger when document is closed; explore using [`Files: watch`](https://developers.google.com/drive/api/v3/reference/files/watch) to approximate onEdit trigger available in Google Sheets).

* Allow customizable syntax for word limit, improve text handling, and add options for character counts in addition to word counts.

* Explore option to publish as an add-on.
