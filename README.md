
Pythons scripts which update, or help with updating, Microsoft Word .docx files.

Use case: you want to generate many Word Documents whose tables and text differ slightly from each other. Data itself already exists in the document and just requires substitution in an intelligent way to preserve the formatting.

The solutions floating around online that I'm aware of are:
- Eric White's Document Assembler (.NET). This requires the template file be organized in a particular way. The tables themselves cannot already contain any data.
- https://github.com/elapouya/python-docx-template. I haven't experimented with this much, however it appears to use a jinja template syntax in the setup.
- python-docx. It's unclear how to use this library to construct template files. Also python-docx doesn't appear to recognize Structured Document Tags (sdts).


The code uploaded to this repository functions in the following way:
1. A Microsoft Word template file is created and formatted through the Word application. There is no special syntax included in the document. Tables and text are tagged through Word's interface. For example, text in the document is highlighted with plain text content controls and tables are captioned/titled. This formatting serves to identify the elements in the document you want to replace.
2. You store the data you wish to use to do the substitutions somewhere, in some kind of intermediate data file. 
3. The script then searches for these tags/captions and does the relevant replacing and re-formatting.

I've tried to document the script itself. Please see the doc strings.
