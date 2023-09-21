# docautoshare

## install
##pip install python-docx
##pip install xlrd
##pip install pandas

pandas is not used in this example, but is used on my automation
...
Pull and run docautomate.py
then run excelautomate.py

each will have inputdoc.docx, inputdoc.docx + input.xls as input, and output.docx, output2.docx as output.

this will create two files, each which..

output - will have text editted with result of 8
on table, it will have new data A1, which is center aligned.

output2 - will have text editted with result from A2 + A3 from Excel
on table, new row is added, and on the (0,0) of table, Data from A1 from Excel is written.

simple example to share for backend implementation.
