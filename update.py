"""
A low-tech standalone script that updates Microsoft Word .docx files by replacing
tagged text in Structured Document Tags (sdts) and Word tables.

Most of the work/formatting is done in the Microsoft Word application itself. Special strings, "tags"
are placed within Plain Text Content Controls (through the Developer Tab/Ribbon). Tables
are captioned by the document creator by accessing their "Properties" menu within Microsoft Word. 
In the menu "Alt Text" of "Table Properties", the table's caption is set in the "Title" text box.



"""


import sys
from lxml import etree
from lxml.etree import QName
from pathlib import Path
import zipfile
import copy
import pandas as pd
from datetime import datetime


# These are some common namespaces/names you see in the underlying document.xml file:
NS = {'w'   : "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "xml" : "http://www.w3.org/XML/1998/namespace"}



class Database:
    """
    This stores the data you want to use to replace your tagged Microsoft Word elements.
    Something that behaves like a dictionary and can be loaded from different sources.
    """
    def __init__(self, data):
        self._data = data
        
    @staticmethod
    def from_json(path):
        import json
        with open(path, 'r') as f:
            json_data = json.load(f)
        return Database(json_data)
    
    def keys(self):
        return [r["name"] for r in self._data]
        
    def items(self):
        return [(r["name"], r['data']) for r in self._data]

    def where(self, predicate):
        subset = [r for r in self._data if predicate(r)]
        return Database(subset)
    


def WElement(localname):
    """Allows us to create w: prefixed lxml Elements"""
    return etree.Element("{" + NS['w'] + "}" + localname, nsmap=NS)

def WSubElement(parent, localname):
    """Allows us to create w: prefixed lxml child Elements"""
    return etree.SubElement(parent, "{" + NS['w'] + "}" + localname, nsmap=NS)
    
def localname(element):
    """The localname of an Element is the part following the prefix, apparently"""
    return etree.QName(element).localname
 

def document_from_path(path):
    """Constructs an lxml Element from a .docx file"""
    with zipfile.ZipFile(path, 'r') as docx:
        xml = docx.read("word/document.xml")
    tree = etree.fromstring(xml)
    documents = tree.xpath("//w:document", namespaces=tree.nsmap)
    assert(len(documents) == 1)
    return documents[0]


def save_document(document, from_path, to_path):
    """Re-builds the .docx zip in from_path using the argument 'document' as the xml for the file word/document.xml. 
    All other files in the zip are identical to the original docx file in from_path"""
    with zipfile.ZipFile(to_path, 'w') as new_docx:
        with zipfile.ZipFile(from_path, 'r') as old_docx:
            for item in old_docx.infolist():
                if item.filename != 'word/document.xml':
                    new_docx.writestr(item, old_docx.read(item.filename))
            new_docx.writestr('word/document.xml', etree.tostring(document, pretty_print=True))


def get_or_add_properties(element):
    """
    XML Elements in Word have property tags with a common pattern. For example, paragraph elements, w:p, have
    property subelements of the form "w:pPr". Likewise, tables in Microsoft Word are "w:tbl" elements whose
    properties are sitting in the subelement "w:tblPr". This function fetches, or creates, the properties 
    subelement.
    """
    tag = localname(element) + "Pr"
    prs = element.xpath("w:" + tag, namespaces=NS)
    if len(prs) == 1:
        return prs[0]
    elif len(prs) == 0:
        pr = WElement(tag)
        element.insert(0, pr)
        pr.text = ''
        return pr
    else:
        raise Exception("This w:{} element has more than one Pr subelement.".format(localname(element)))


def copy_properties(from_element, to_element):
    """
    Copies the properties from from_element to to_element by manipulating the lxml (sub)Elements.
    """
    pr_other = get_or_add_properties(from_element)
    pr       = get_or_add_properties(to_element)
    to_element.replace(pr, copy.deepcopy(pr_other))
    
    

def copy_tc_properties(from_tc, to_tc):
    """
    A tc element usually has a nested structure tc/p/r and
    each of these subelements need to have their properties copied.
    """ 
    copy_properties(from_tc, to_tc)
    from_p = from_tc.xpath("w:p", namespaces=NS)
    to_p = to_tc.xpath("w:p", namespaces=NS)
    copy_properties(from_p[0], to_p[0])
    from_r = from_p[0].xpath("w:r", namespaces=NS)
    to_r = to_p[0].xpath("w:r", namespaces=NS)
    copy_properties(from_r[0], to_r[0])


def copy_table_properties(from_tbl, to_tbl):
    """
    The formatting properties of a table are specified in the nested subelements 
    of the w:tblPr tag. This function loops through the relevant sub-Elements and copies
    their properties to to_tbl.
    
    We assume the original table, from_tbl, has at least two rows (one that represents
    the formatting of the headers and another that represents the formatting of the data).
    """
    copy_properties(from_tbl, to_tbl)
    row_count = len(from_tbl.xpath("w:tr", namespaces=NS))
    header_row   = from_tbl.xpath("w:tr[1]", namespaces=NS)
    header_cells = header_row[0].xpath("w:tc", namespaces=NS)
    data_row     = from_tbl.xpath("w:tr[{}]".format(3 if row_count >= 3 else 2), namespaces=NS)
    data_cells   = data_row[0].xpath("w:tc", namespaces=NS)
    #assert(len(header_cells) == len(data_cells))
    for i, row in enumerate(to_tbl.xpath("w:tr", namespaces=NS)):
        cells = row.xpath("w:tc", namespaces=NS)
        copy_from = header_row if i==0 else data_row
        copy_properties(copy_from[0], row)
        for j in range(len(header_cells)):
            copy_from = header_cells if i==0 else data_cells
            copy_tc_properties(copy_from[j], cells[j])



def update_table(table, replacement_table):
    """
    Updates table by 
    1) replacing its text with the content in the replacement, and
    2) copying the relevant (nested) Pr data in the original. 
    """
    copy_table_properties(table, replacement_table)
    new_rows = copy.deepcopy(replacement_table.xpath("w:tr", namespaces=NS))
    rows = table.xpath("w:tr", namespaces=NS)
    for row in rows:
        table.remove(row)
    for row in new_rows:
        table.append(row)
     

def table_from_html(html):
    """
    Helper function. Pandas DataFrames are easily converted to HTML.
    We then use the HTML and convert it to the appropriate XML Element
    to be used within a Microsoft Word Document.
    """
    table = etree.fromstring(html)
    tbl = WElement("tbl")
    for row in table.xpath("//table[1]/descendant::tr"):
        cells = row.xpath("th|td")
        if len(cells) == 0:
            continue
        tr = WSubElement(tbl, "tr")
        for cell in cells:
            tc = WSubElement(tr, "tc")
            p  = WSubElement(tc, "p")
            r  = WSubElement(p , "r")
            t  = WSubElement(r , "t")
            t.text = cell.text
    return tbl



def get_text_element(run_element):
    """
    Returns the nested w:t element under a run w:r element.
    """
    text_elements = run_element.xpath("w:t", namespaces=NS)
    assert(len(text_elements) == 1)
    return text_elements[0]


def replace_sdt_text(sdt, text):
    """
    Replaces the text in the w:sdt Element. Microsoft Word separates
    text into different "runs". A "run" is just a container for text
    with identical formatting properties.
    
    This function first merges all runs into the first run element 
    and then writes the text of that single run element.
    """
    runs = sdt.xpath(".//w:r", namespaces=NS)
    assert(len(runs) > 0)
    first = get_text_element(runs[0])
    first.text = '' if first.text is None else first.text
    first_parent = runs[0].getparent()
    for i, run in enumerate(runs):
        if i > 0:
            assert(run.getparent() is first_parent)
            first.text = first.text + get_text_element(run).text
            first_parent.remove(run)
    first.text = text


def get_sdt_by_tag(document, tag_value):
    """
    Finds w:sdt elements which have a w:tag element with w:val == tag_value
    """
    return document.xpath(".//w:sdt/w:sdtPr/w:tag[@w:val='{}']/../..".format(tag_value), namespaces=NS)

def get_table_by_caption(document, caption):
    """
    Finds w:tbl elements captioned (or Titled) with caption.
    """
    tbls = document.xpath(".//w:tbl/w:tblPr/w:tblCaption[@w:val='{}']/../..".format(caption), namespaces=NS)
    if len(tbls) == 0:
        raise Exception("There isn't a table captioned \"{}\"".format(caption))
    elif len(tbls) > 1:
        raise Exception("There's more than one table captioned \"{}\"".format(caption))
    else:
        return tbls[0]


def get_output_path(identifier):
    """
    Some kind of naming scheme for the output files.
    """
    timestamp = str(datetime.now()).replace(" ", "_").replace(":", "_").split(".")[0]
    #new_filename = "UPDATED_" + path.stem + "_" + timestamp + path.suffix
    #return path.parents[0] / new_filename
    return Path("Updated_{}_{}.docx".format(identifier, timestamp))
    
 
def main():
    
    if len(sys.argv) < 3:
        print("You didn't use the program correctly.")
        print("Usage: <project data file> <template docx file>")
        print("Exiting...")
        sys.exit(1)

    database_path = Path(sys.argv[1])
    filename = Path(sys.argv[2])
    
    db = Database.from_json(database_path)
    doc = document_from_path(filename)

    for src in ["Worksheet1", "Worksheet2"]:
        subdb = db.where(lambda r: r["worksheet"]==src)
        for key, val in subdb.items():
            if key.startswith("TABLE_"):
                caption = key[6:]
                print("Updating table with caption \"{}\"".format(caption))
                tbl = get_table_by_caption(doc, caption)
                new_tbl = pd.DataFrame(val)
                new_tbl = new_tbl.to_html(index=False)
                new_tbl = table_from_html(new_tbl)
                update_table(tbl, new_tbl)
            else:
                print("Replacing tags {} to have text \"{}\"".format(key, val))
                sdts = get_sdt_by_tag(doc, key)
                if len(sdts) == 0:
                    print("WARNING: no tags named \"{}\" (no replacements made)".format(key))
                for sdt in sdts:
                    replace_sdt_text(sdt, val)
        
        outfile = get_output_path(src)
        print("Saving updated docx file to \"{}\"".format(outfile.stem))
        save_document(doc, filename, outfile)
        print("Success!!")


if __name__ == "__main__":
    main()
    




