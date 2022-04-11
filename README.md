# VBA-Password-Remover
A work in progress project. This tool removes passwords at a workbook and worksheet level. Currently working on a method to remove project level passwords.

Passwords at workbooks and worksheet levels are removed by;
1. Renaming the .xlsx/.xlsm/.xlsb file as a .zip file
2. Opening the zip file and working with the xml structure of the workbook, the password tags are removed.
This section is working

Removing passwords at a project level is not working yet but I have proven the method. The file needs manipulating at a byte level so a couple of key bytes can be removed. If I use a HexEditor outside of Excel this is a simple search and replace exercise.
