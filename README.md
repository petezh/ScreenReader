# ScreenReader

This is a VBA macro to convert Verbatim speech documents to a screen-reader accessible format. Please read the usage notes!

### Usage Notes

1. The macro is in development and may cause issues on non-Windows systems. Please report any issues to petez@berkeley.edu.

2. The macro will convert the current file that is open. It is recommended that users open a new Word document and paste in the content they wish to convert before running the macro.

3. Should a user accidentally convert a document, they can simply undo the operation by pressing Control+Z twice.

4. Make sure the following conditions are true for documents (some are less important than others):
- Tags and analytics are all marked with "tag"
- Card names are marked with either "tag" or "cite"
- Parts of cards to read are delineated with highlighting
- Section headers that should be kept in the doc are styled as "pocket", "hat", or "block"

### Macro Installation

1. Download the file titled 'CopyForScreenWriter.txt', in the repository.

2. Open the file in any file editor (e.g. Notepad, TextEdit, Notepad++).

3. Copy the all of the code with Control-A and Control-C.

4. Open any Word document.

5. Navigate to the View tab and press on 'Macros', or 'View Macros.'

6. In the 'Macro name' box, type 'CopyForScreenReader.'

7. Select 'Create.' This should bring you to the VBA editor.

8. Paste the code. The code should be situated between the lines 'Sub CopyForScreenReader()' and 'End Sub.'

9. Save the macro (with Control-S) and close out of the VBA editor.

10. To verify that your macro has been installed, press 'Macros' once again and scroll through the list of available Macros.

11. Open any document that you wish to convert. Open 'Macros', select 'CopyForScreenReader,' and press 'Run.'

For more information on steps 1-9, please visit:
http://www.techtoolsforwriters.com/how-to-add-a-macro-to-word/

### Shortcut Installation

1. Press the 'File' tab and select 'Options' in the bottom left corner.

2. Press 'Customize Ribbons.'

3. Next to 'Keyboard Shortcuts:', press 'Customize'.

4. Scroll to the bottom of the list of 'Categories' and press 'Macros.'

5. In the right hand dropdown, select 'CopyForScreenReader.'

6. In the box titled 'Press new shortcut key,' press the desired shortcut combination key, e.g. Control+=.

7. Ensure that changes are saved in the 'Normal' document template. This ensures the shortcut applies to all future documents.

8. Create a new document or restart word. Paste in debate content, and test the shortcut!

For more information on steps 1-7, please visit:
https://wordribbon.tips.net/T008058_Assigning_a_Macro_to_a_Shortcut_Key.html

Thank you Xavier for developing the tool.




