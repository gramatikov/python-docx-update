
import win32com.client as win32
from pathlib import Path
import sys
import json

DEFAULT_OUTPUT_DIRECTORY = Path("./ProjectDatabase")


def dump_named_ranges(wb, path):
    data = []
    for name in wb.Names:
        record = {}
        r = name.RefersToRange
        record['address'] = str(r.Address)
        tokens = str(name.Name).split('!')
        i = 1 if len(tokens) == 2 else 0
        record['name'] = tokens[i]
        record['worksheet'] = str(r.Worksheet.Name)
        record['rows'] = r.Rows.Count
        record['columns'] = r.Columns.Count
        if r.Rows.Count == 1 and r.Columns.Count == 1:
            record['data'] = str(r.Rows[1].Cells[1].Text)
        else:
            table = {}
            headers = [str(cell.Text) for cell in r.Rows[1].Cells]           
            for i, row in enumerate(r.Rows):
                for j, cell in enumerate(row.Cells):
                    if i==0:
                        table[headers[j]] = []
                    else:
                        table[headers[j]].append(str(cell.Text))
            record['data'] = table
        data.append(record)
        
    with open(path / '{}.json'.format(str(wb.Name).split(".")[0]), 'w') as f:
        json.dump(data, f)

  
def main():
    
    if len(sys.argv) == 1:
        print("Error: must specify the Excel file name (nothing was done)")
        print("Usage: <excel file>")
        sys.exit(1)

    filename = Path(sys.argv[1])
    
    app = win32.gencache.EnsureDispatch('Excel.Application')
    wb = app.Workbooks.Open(filename.resolve())
    
    DEFAULT_OUTPUT_DIRECTORY.mkdir(exist_ok=True)
    dump_named_ranges(wb, DEFAULT_OUTPUT_DIRECTORY)

    wb.Close(SaveChanges=False)
    app.Quit()
    print("Success!")

    
    
main()
