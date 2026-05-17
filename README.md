
Pythons scripts which update Microsoft Word .docx files.

Use case: you want to generate many Word Documents whose tables and text differ slightly from each other. 

There are two solutions floating around online to accomplish this:
- https://github.com/elapouya/python-docx-template
- Eric White's Document Assembler (.NET)

The code uploaded to this repository functions in the following way:
1. A Microsoft Word template file is created and formatted through the Word application. Plain Text Content Controls are added to text which you want to replace. Tables are captioned/titled through the application as well.
2. The script then searches for these tags/captions and does the relevant replacing and re-formatting.

