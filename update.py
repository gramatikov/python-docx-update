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
from datetime import datetime


# These are some common namespaces/names you see in the underlying document.xml file:
NS = {'w'   : "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "xml" : "http://www.w3.org/XML/1998/namespace"}



class Database:
    """
    My interface for interacting with the data used to do the replacing.
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
    pr = get_or_add_properties(to_element)
    to_element.replace(pr, copy.deepcopy(pr_other))
    
    

def copy_tc_properties(from_tc, to_tc):
    """
    A tc element usually has a nested structure tc/p/r and
    each of these subelements need to have their properties copied.
    """ 
    copy_properties(from_tc, to_tc)
    from_p = from_tc.xpath("w:p", namespaces=NS)
    # if the original w:tc element has a paragraph w:p, copy from it
    if len(from_p) > 0:
        to_p = to_tc.xpath("w:p", namespaces=NS)
        copy_properties(from_p[0], to_p[0])
        from_r = from_p[0].xpath("w:r", namespaces=NS)
        # if the original w:tc/w:p element has a run w:r, copy from it
        if len(from_r) > 0:
            to_r = to_p[0].xpath("w:r", namespaces=NS)
            copy_properties(from_r[0], to_r[0])




def copy_table_properties(from_tbl, to_tbl):
    """
    The formatting properties of a table are specified in the nested subelements 
    of the w:tbl tag. This function loops through the relevant subelements and copies
    their properties to to_tbl.
    
    We assume the original table, from_tbl, has at least two rows 
    """
    original_rows = from_tbl.xpath("w:tr", namespaces=NS)
    new_rows = to_tbl.xpath("w:tr", namespaces=NS)
    cell_sequence = [len(tr.xpath("w:tc", namespaces=NS)) for tr in original_rows]
    copy_properties(from_tbl, to_tbl)
    # if the original table only has two rows, then the second row, at index 1, will be the first proper row
    # otherwise, the first proper row will be the third row at index 2. This helps the formatting of the table
    first_proper_row = 2 if len(original_rows) >= 3 else 1
    for i, new_row in enumerate(to_tbl.xpath("w:tr", namespaces=NS)):
        new_cells = new_row.xpath("w:tc", namespaces=NS)
        # we find the index of the row whose formatting the new_row will copy from
        # if we see a ValueError, there is no proper row in the original table having the same number of cells
        copy_from = 0 if i==0 else (-1 if i==len(new_rows)-1 else cell_sequence.index(len(new_cells), first_proper_row))
        copy_from = original_rows[copy_from]
        copy_properties(copy_from, new_row)
        for j in range(len(new_cells)):
            # copy_from is a row which will have the same number of cells as new_cells
            copy_from_cells = copy_from.xpath("w:tc", namespaces=NS)
            copy_tc_properties(copy_from_cells[j], new_cells[j])



def update_table(table, replacement_table):
    """
    Updates the argument table by 
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
     

def table_from_list_of_lists(lst):
    """
    Constructs the w:tbl element.
    """
    tbl = WElement("tbl")
    for row in lst:
        tr = WSubElement(tbl, "tr")
        for cell in row:
            tc = WSubElement(tr, "tc")
            p  = WSubElement(tc, "p")
            r  = WSubElement(p , "r")
            t  = WSubElement(r , "t")
            t.text = cell
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
    and then writes the text of that single run element. I need
    to update this. It just has to delete all i>0 runs.
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
    return document.xpath(".//w:tbl/w:tblPr/w:tblCaption[@w:val='{}']/../..".format(caption), namespaces=NS)


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

    # some kind of loop to iterate over each "report instance"
    for src in ["Report1", "Report2"]:
        subdb = db.where(lambda r: r["worksheet"]==src)
        # for each key-value pair in the report data
        for key, val in subdb.items():
            # update each tagged/captioned element. I used a simple naming scheme to distinguish tables from text
            if key.startswith("TABLE_"):
                caption = key[6:]
                print("Updating table with caption \"{}\"".format(caption))
                tbls = get_table_by_caption(doc, caption)
                if len(tbls) == 0:
                    print("WARNING: no table captioned \"{}\" (no replacements made)".format(caption))
                    continue
                new_tbl = table_from_list_of_lists(val)
                for tbl in tbls:
                    update_table(tbl, new_tbl)
            else:
                print("Replacing tags {} to have text \"{}\"".format(key, val))
                sdts = get_sdt_by_tag(doc, key)
                if len(sdts) == 0:
                    print("WARNING: no tags named \"{}\" (no replacements made)".format(key))
                for sdt in sdts:
                    replace_sdt_text(sdt, val)
        
        outfile = get_output_path(filename.stem + "_" + src)
        print("Saving updated docx file to \"{}\"".format(outfile.stem))
        save_document(doc, filename, outfile)
        print("Success!!")


if __name__ == "__main__":
    main()
    




