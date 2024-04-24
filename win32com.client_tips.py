import win32com.client

# Create an instance of the Word application
word_app = win32com.client.Dispatch("Word.Application")
word_app.Visible = True  # Make Word visible (optional)

# Open an existing Word document
doc_path = "example.docx"  # Replace with the path to your Word document
doc = word_app.Documents.Open(doc_path)

# Set the selection range within the document
selection = word_app.Selection
selection.Start = 0  # Start from the beginning of the document
selection.End = selection.Document.Characters.Count  # End at the end of the document

# Move the selection down to the end of the document (equivalent to selecting the entire story)
selection.MoveDown(Unit=6, Count=1)  # Unit=6 corresponds to wdStory

# Close the document and quit Word application
doc.Close()
word_app.Quit()

# possible common constants and their corresponding integer values, from chatgpt, they not all correct, test before use
# wd refers word
"""
wdStory: Represents a story, which generally corresponds to the entire contents of the document. Integer value: -4.

wdCharacter: Represents a character. Integer value: 1.
wdWord: Represents a word. Integer value: 2.
wdParagraph: Represents a paragraph. Integer value: 4.
wdLine: Represents a line of text. Integer value: 5.
wdSection: Represents a section. Integer value: 6.
wdScreen: Represents the screen. Integer value: 7.
wdColumn: Represents a column. Integer value: 9.
wdRow: Represents a row. Integer value: 10.
wdWindow: Represents a window. Integer value: 11.
wdCell: Represents a cell. Integer value: 12.
wdCharacterFormatting: Represents character formatting. Integer value: 13.
wdParagraphFormatting: Represents paragraph formatting. Integer value: 14.
wdTableFormatting: Represents table formatting. Integer value: 15.
wdList: Represents a list. Integer value: 16.
wdDocument: Represents a document. Integer value: 17.
wdFrame: Represents a frame. Integer value: 18.
wdTable: Represents a table. Integer value: 19.
"""
