EveCentralImports
=================

Pulls current prices into a spreadsheet

Cell column A is for the item name (please use the EXACT item name, need the correct name for eve central to find it)

The program takes input.xml and outputs a output.xml with all cells copied over.

Anything in Cell C in the same row as an item name in column A will be overwritten with eve-central data

Blacklist.txt has names of headers that might be in use in column A.
If there are headers in use that are not in the blacklist.txt the program will stop when it gets to a name that is not an item.

The excel sheets must be .xml format, the libraries I'm using to read and write to the excel sheet don't support .xmlx
