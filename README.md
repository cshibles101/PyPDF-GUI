# PyPDF-GUI
Python Tkinter project that allows a user to choose files in a graphical user interface and either combine it with another pdf file or split it into multiple smaller pdf files

# Combining PDFs
Users are presented with 3 Radio Buttons to choose from and Combine is the default. This presents the user with 3 Entry fields that when double clicked, opens a save as dialog to allow the user to select which pdf documents in their files they wish to combine into one and also a file path to save the combined document to with their document name of choice. Open choosing a value for all 3, the user can click Go and if the combination was successful, all 3 fields are emptied. If they are not successful for whatever reason, a red message label appears above the Go button that hints at what the issue may have been.

# Splitting PDFs
The second Radio Button presents the user with 2 Entry fields. The first is for a file path that allows a double click to choose the file from the file explorer. The second allows for 3 types of entries that tells the program how to split the file. Entering a single number will split the file at that page in the pdf (the number enterred being the first page of the second document after the split). This value cannot be 1 or greater than the length of the pdf document. Entering multiple numbers separated by commas (,) will cause the PDF to split at each of those page numbers. Entering a number followed by an asterisk (\*) will cause there to be an interval split which tells the program the user wants to split the document at every *nth* pages when entering *n\**. These formats are very specific and the program uses Regex to ensure the format is correct and if it is not correct, the message label above the Go Button will warn the user. Upon successfully splitting the document, the Entry fields will be cleared and the new PDFs will be in the same directory as the original. The naming convention for every new document is just the name of the original document followed by an underscore and the number of the first page in the new document.

# Watermarking PDFS
The third Radio button presents the user with 2 Entry fields. Both fields allow the user to double click to open the file explorer to choose PDF files to be used. The first is the main PDF and the second is for the watermark image pdf. Once a file is selected for both, the user can click Go to process the watermark. Watermarked files are saved with the same name as the main PDF file but with '_watermarked' after the file name and before the extension. 
