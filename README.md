# anon-docs-py

small program to delete whichever words from a csv list in a docx files
as its made for specific csv-lists from a specific source, key are hardcoded. Might eventually change to dynamic choosing of keys in guy.

- ~~todo: add source document through gui~~
- todo: add preview after cleaning before saving
- ~~todo: add save as name+path and this logic to gui~~
- import data from csv
  - ~~todo: add file selection in gui~~
  - todo: add parameter specification in gui
  - todo: dynamic preview???
    
eventually refactor and prettify gui

To Use:
- Prepare docx:
  - search and replace all soft returns (^l) with carriage returns (^p)
- Prepare CSV
  - Make sure all entries to be deleted from the document are present
  - names with hyphen will be considered one name, names with spaces will be considered two

All additional formatting will be lost