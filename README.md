# box2docx
A little program to convert Box notes to Word docx files. Functionality is limited to the new Box Note format. Files created before Aug. 2022 are not supported and will not be converted.

I threw this together since the tools that others have provided on GitHub don't convert all the different styles and items in the Box note, or rely on having an access token through the API. My other goal was to create output that closely resembles the style and look of the original Box note.

The following items should convert correctly:
- Headings, Body text
- Bold, Italic, Underline, Strikethrough, and combinations of
- Font sizes
- Font colors
- Highlight colors (although colors had to be mapped to the few other highlight colors provided by Word)
- Text alignments
- Checklist (Word doesn't support checkboxes. Instead we draw either a ☑ or a ☐ based on if the box was checked)
- Number lists 
- Bullet lists
- Images (Images are scaled for the Word Document, but can be manually resized)
- Tables (including merged cells)
- Divider lines
- Call outs
- Code blocks
- Block quotes

The following are not supported:
- Comments
- Embedded files

Since default words styles only allow for 3-level list, and I have lists with more levels. I instead chose to write out lists manually by indenting and putting numbering the lists myself.


This program has currently only been tested on OS X, but may work with Windows as well.

## Requirements
- Box Drive: This program requires that you have Box Drive installed on your machine. It relies on the directories and files Box Drive creates to read Box Notes and convert Box Notes.
- A python 3.10 environment.

## Setup
- Clone the repository.
- Set up an python 3.10 environment. (Other versions may work, but this was tested in 3.10.)
- In the project root, install the requirements `pip install -r requirements.txt'

## Usage
`usage: box2docx.py [-h] [--format {docx,md,html}] [--recursive] [--update_legacy_boxnotes] [--dry-run] [--debug] path`

The box notes must be in folder created by Box Drive. Default locations are:
- On Mac: `~/Library/CloudStorage/Box-Box/`
- On Windows: `C:\Users\<USERNAME>\Box\`

On Mac your command should look something like this to convert one file:

`python box2docx.py ~/Library/CloudStorage/Box-Box/path/to/box/folder/file.boxnote`

The program will create a .docx file to the same directory with the same name as the boxnote. **It does not check if a file with that name already exists!** So if you have a .docx and .boxnote with the same name in the directory, the .docx will be overwritten.

You can also use the terminals file expansion to easily convert all boxnote files in a single folder:

`python box2docx.py ~/Library/CloudStorage/Box-Box/path/to/box/folder/*.boxnote`

## Considerations
Sometimes it takes a while for a Box Note to be downloaded to your machine. The tool does include an additional 10 second wait, but that may not be enough. You may need to rerun the tool after the file is fully loaded if it doens't work the first time.