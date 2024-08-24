# Text Extractor with Tkinter UI

This Python script provides a simple graphical user interface (GUI) built with Tkinter that allows users to extract specific rows from a Word document (`.docx`) based on a numbering pattern (e.g., `1#h`, `2#h`, etc.). The extracted content can be saved as either a text file (`.txt`) or another Word document (`.docx`), preserving formatting such as bold, italic, underline, font size, and font color.

## Features

- **File Selection**: Users can select a Word document from their file system.
- **Output Format**: Users can choose to save the extracted content as either a text file or a Word document.
- **Formatting Preservation**: When saving as a Word document, the script preserves text formatting, including colors, bold, italic, underline, and font sizes.
- **Simple GUI**: The user interface is straight forward and easy to use, with options for file selection and output format.

## Prerequisites

Before running the script, ensure you have the following Python libraries installed:

- `tkinter` (usually included with Python)
- `python-docx`
- `python 3.x`
- `pip`

You can install `python-docx` using pip:

```bash
pip install python-docx
``` 
## How to Use

Run the Script: Execute the script in a Python environment that supports Tkinter.

```bash
python3 extract.py
```

Select the File: Click the "Browse..." button to select the Word document (.docx) from which you want to extract content.

Choose Output Format: Use the radio buttons to select the desired output format:
Text File (.txt): The extracted content will be saved as a plain text file.
Word Document (.docx): The extracted content will be saved as a Word document, with formatting preserved.

Run Extraction: Click the "Run Extraction" button to start the extraction process. You will be prompted to choose where to save the output file.

Success Message: After the extraction is complete, a message will appear confirming that the data has been saved to the specified location.




