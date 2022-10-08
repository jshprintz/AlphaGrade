## AlphaGrade

#### AlphaGrade is a series of macros written in VBA designed to facilitate and standardize the grading process for QAs scoring transcriptionists.

---

### Install Instructions - (Windows)
1. Open Windows Explorer/File Exporer
2. Copy/Paste the following into the address bar:
    - `%Appdata%/Microsoft/word`
    - If there isn't already a folder named *STARTUP*, create one
3. Open the *STARTUP* folder
4. Copy AlphaGrade.dotm file into the *STARTUP* folder.
5. Restart Word

### Install Instructions - (Mac)
1. Double-click on the ALphagrade.dotm file
2. When prompted, click *Enable Macros*
3. Click on the *View* tab
4. Click on *Macros* and then the *Organizer* button
5. On the left, where it says, "Macro Project Items available in:", click the dropdown menu and select AlphaGrade.dotm
6. Click on each one and then click the copy button to transfer them to Normal.dotm
    - NOTE: if the file name changes for whateevr reason, click rename and change the file back to the original name displayed on the left
7. Click *Close* and restart Word

---

### Using Macros
On the *View* tabe, you'll see a *Macros* button. By clicking on the button, you'll see the available macros. You can double click or click run to use them.

#### QA Comments
QA Comments serves as a way to both uniform the grading comments, so that all QAs will be delivering consistent feedback to TRs in a consistent manner, and to allow for quick grading based upon the grading calculator. When running the program, the default is at minor mishear, but the user can select from any of the content or formatting errors. 

#### QA Checklist
QA Checklist works together with QA Comments. It also offers several standalone features. 

#### Tag Count
Tag Count will tell you the total number of inaudibles, guesses, crosstalks, foreign tags, and research tags in any file.

#### Pre Grade
The TRs are given a macro that serves to clean up their work before submitting. If they fail to use that macro, Pre Grade will essentially do all of those processes and a few more. 
- Eliminates instances of two spaces that start a sentence
- Replaces straight quotes with curly quotes
- Doesn’t change the user’s specified settings
- Searches for incorrectly formatted timestamps
- Anonymizes your username
- Searches for incorrect punctuation
- Replaces the following words
- Runs bug fixes in the background to correct a common spacing issue that Word creates caused by incorrect borders 
- Stores the initial number of inaudibles and guess tags

Since the TRs are expected to run this macro prior to submitting, when they don’t run it, all changes will be marked against their final score. After running Pre Grade, a dialog box will appear on the right detailing if any points have been deducted. Since the TRs are expected to run their macro, the large majority of the time their score will stay at 100. 

#### Post Grade
After you are finished reviewing the transcript, Post Grade will calculate all of the scoring for you. It will start with the Pre Grade score and then it will subtract all of the major errors and minor errors. It will also work off the following grade system for determining cleared inaudibles/guesses:
    1-3 = 0
    4-6 = -0.25
    over 7 = -0.5
(Example: Eight cleared inaudibles would be a deduction of -1.75. Zero deducted for the first three, -0.25 each for inaudibles four, five, and six, and then -0.5 each for seven and eight.)

#### Client Ready
The Client Ready macro is designed to prepare a document to be ready for submission to the client. Client Ready is to be used after the split is uploaded and graded. Client Ready performs a number of functions:
- Accepts all edits and turns off trackchanges
- Deletes all comments
- Runs spellcheck
- Reverts username back to original setting
- Ensures that there are no research tags left in the document
- If there are, it alerts you that there are still research tags and that the file is not ready yet.
Once the file is ready, a dialog box will appear stating, "This document is client ready."
