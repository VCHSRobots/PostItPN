# spn.py -- Sticky Part Numbers
# dlb, Dec 2023
#
# To run this code, you must install the following:
#
#  pip install pandas
#  pip install openpyxl
#  pip install xlrd
#  pip install reportlab
#

Instructions = '''
This is a program for creating Sticky Notes with Part Numbers on them.
Usage:  python spn.py input-file output-file [final|template]

The input file should be an excel (xls, xlsx) file with one sheet, or multiple sheets
with one of the sheets named 'bom'.  The sheet should have the following columns:
   "Part Number"    -- Required, data in the form "XNN-AA-0000"
   "Quantity"       -- Required, a number
   "Sub Number"     -- Optional, such as "mirror", or "Confg 1"
   "Description"    -- Optional, very short, such as "Left Bracket"
   "Material"       -- Optional, such as "PETG", "6061 Aluminum", "Plywood"
   "Stock"          -- Optional, such as "1x1x0.125 AL Sq TUBE"
   "Guild"          -- Optional, such as "CNC", "MACH", "BUILD", "CAD"
   "Machine"        -- Optional, such as "THOR", "TONY", "MILL", "LATH", "DORINGER", "HULK"
   "Group"          -- Optional, one or two letters, such as "A"
   "Designer"       -- Optional, name of designer, such as "IShiraki"

The output file is a PDF that prints on letter sized pages, and is formatted for
2x2 inch post-it notes.  

The parameter after the output-file can be either "final" or "template".  If left unspecified,
then "template" is assumed. When "template" is used, an outline is placed around each
post-it note, whereas there is no outline when "final" is used.

How to use: after preparing your BOM in a spread sheet, run this program and look
for any error messages.  Fix the spread sheet until there are no errors.  Then print
the resulting pdf with "template".  The output will be a template on which to put post-it
notes over each outlined square.  Reload the paper with the post-it notes in the
printer, and re-run the program with "final".  Print the final pdf on the loaded
paper with the post-it notes.

(Version 1.0, Dec 2023)
'''

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch 
from reportlab.lib.pagesizes import letter
import pandas as pd
import numpy
import sys
import os
import math 


do_outline = False
do_developer = False
note_size = (2.0, 2.0)                   # The size of the sticky note.
note_margin = 0.05                       # The distance between the template line and the note      
page_margin = (0.75, 0.75, 0.5, 0.5)     # The printer's page margins: top, left, bottom, right 
grid = (3,4)                             # Tells how many notes on a page: across by up-down

def get_value(v, maxc, vname, pn, iline):
    ''' Returns a string for the input data.'''
    d = ""
    if type(v) is str: d = v
    if type(v) is int: d = str(v) 
    if type(v) is numpy.int64: d = str(v)
    if type(v) is numpy.float64 or type(v) is numpy.float32 or type(v) is float:
        if math.isnan(v): d = "" 
        else:
            n = int(v + 0.499)
            d = str(n)
    if len(d) > maxc:
        print("%s is too long for %s in line %d.  Value truncated!" % (vname, pn, iline))
        d = d[0:maxc]
    return d

def rddata(filename):
    '''Reads the excel file, returns a list of records, where each record is a
    dict of name-value pairs.  None is returned on error.'''
    try:
        all_sheets = pd.read_excel(filename, sheet_name=None)
    except:
        print("Unable to read %s." % filename)
        return None
    if len(all_sheets) > 1:
        if "bom" not in all_sheets.keys():
            print("Too many sheets (%d), and sheet named 'bom' not found." % len(all_sheets))
            return None
        df = all_sheets['bom']
    elif len(all_sheets) == 1:
        k = list(all_sheets.keys())[0]
        df = all_sheets[k]       
    else:
        print("No sheets found.")
        return None 
    if len(df.columns) <= 0:
        print("No columns found.")
        return None
    if "Part Number" not in df.columns:
        print("Part Number column not found.")
        return None
    if "Quantity" not in df.columns:
        print("Quantity column not found.")
        return None
    PartNumber  = df.get("Part Number")
    Quantity     = df.get("Quantity")
    n = len(PartNumber)
    SubNumber   = [""] * n
    Description = [""] * n
    Material    = [""] * n
    Stock       = [""] * n
    Guild       = [""] * n
    Machine     = [""] * n
    Group       = [""] * n
    Designer    = [""] * n
    if "Sub Number" in df.columns: SubNumber = df.get("Sub Number")
    if "Description" in df.columns: Description = df.get("Description")
    if "Material" in df.columns: Material = df.get("Material")
    if "Stock" in df.columns: Stock = df.get("Stock")
    if "Guild" in df.columns: Guild = df.get("Guild")
    if "Machine" in df.columns: Machine = df.get("Machine")
    if "Group" in df.columns: Group = df.get("Group")
    if "Designer" in df.columns: Designer = df.get("Designer")
    notes = [] 
    for i, pnx in enumerate(PartNumber): 
        iline = i + 2
        pn = get_value(pnx, 50, "Part Number", iline, "??")
        if pn == "": continue       # Skip blank lines
        if len(pn) > 11:
            print("Line %d does not appear to be a part number." % iline)
            pn = pn[0:11]
        if pn[3:4] != '-' or pn[6:7] != '-':
            print("Line %d does not appear to be a part number." % iline)
        info = {"Part Number" : pn}
        try:
            qstr = get_value(Quantity[i], 4, "Quantity", pn, iline)
            q = int(qstr)
        except: 
            q = 0 
        if q <= 0 or q > 999: 
            print("Invalid quanity found for %s in line %d. Using ONE." % (pn, iline))
            q = 1
        info["Quantity"] = q
        info["Sub Number"] = get_value(SubNumber[i], 7, "Sub Number", pn, iline)
        info["Description"] = get_value(Description[i], 25, "Description", pn, iline)
        info["Material"] = get_value(Material[i], 16, "Material", pn, iline)
        info["Stock"] = get_value(Stock[i], 22, "Stock", pn, iline)
        info["Guild"] = get_value(Guild[i], 5, "Guild", pn, iline)
        info["Machine"] = get_value(Machine[i], 25, "Machine", pn, iline)
        info["Group"] = get_value(Group[i], 4, "Group", pn, iline) 
        info["Designer"] = get_value(Designer[i], 16, "Designer", pn, iline)
        notes.append(info)
    return notes

def draw_note(data, location, pdf_canvas):
    ''' Draws one note on the pdf, at the location.  Data is a dict of
    info, location is a 2-tuple specifing the lower left in canvas units,
    and pdf_canvas is the canvas to draw into.'''
    x0, y0 = location[0] * inch, location[1] * inch
    xbox, ybox = note_size  
    xbox_with_margin, ybox_with_margin = xbox + 2*note_margin, ybox + 2*note_margin
    pdf_canvas.setStrokeColorRGB(0.0, 0.0, 0.0)
    pdf_canvas.setLineWidth(2)
    if do_outline:
        x1, y1 = x0 + xbox_with_margin*inch, y0 + ybox_with_margin*inch
        pdf_canvas.line(x0, y0, x0, y1)
        pdf_canvas.line(x0, y1, x1, y1)
        pdf_canvas.line(x1, y1, x1, y0)
        pdf_canvas.line(x1, y0, x0, y0)
        x0, y0 = x0 + note_margin * inch, y0 + note_margin * inch
    if do_developer:
        pdf_canvas.setLineWidth(1)
        x1, y1 = x0 + xbox*inch, y0 + ybox*inch
        pdf_canvas.line(x0, y0, x0, y1)
        pdf_canvas.line(x0, y1, x1, y1)
        pdf_canvas.line(x1, y1, x1, y0)
        pdf_canvas.line(x1, y0, x0, y0)
    x, y = x0, y0 + (ybox - 0.275)*inch
    xc = x + 0.5*xbox*inch
    xm = 0.125 * inch
    xr = x + xbox*inch 
    pdf_canvas.setFont("Courier-Bold", 18.0)
    pdf_canvas.drawCentredString(xc, y, data["Part Number"])
    y -= 0.3*inch
    s = "x%d" % data["Quantity"]
    pdf_canvas.setFont("Courier-Bold", 22.0)
    pdf_canvas.drawCentredString(xc, y, s)
    s = data["Description"]
    if len(s) > 30: s = s[0:30]
    pdf_canvas.setFont("Times-Roman", 12.0)
    y -= 0.35*inch
    pdf_canvas.drawCentredString(xc, y, s)
    if data["Stock"] != "":
        s = data["Stock"]
    elif data["Material"] != "":
        s = "Material: %s" + data["Material"]
    else: s = ""
    y -= 0.35*inch
    if s != "":
        pdf_canvas.setFont("Helvetica-Bold", 12.0)
        pdf_canvas.drawString(x + xm, y, s)
    ymachine = y0 + 0.5  * inch 
    ydesigner = y0 + 0.22 * inch
    yguild   = y0 + 0.05 * inch
    ygroup   = y0 + 0.05 * inch
    ysubnum  = y0 + (ybox - 0.43) * inch
    pdf_canvas.setFont("Helvetica", 12.0)
    if data["Machine"] != "":  pdf_canvas.drawString(x + xm, ymachine, data["Machine"])
    if data["Guild"] != "":    pdf_canvas.drawString(x + xm, yguild, data["Guild"])
    if data["Group"] != "":    pdf_canvas.drawRightString(xr - xm, ygroup, data["Group"])
    if data["Designer"] != "": pdf_canvas.drawRightString(xr - xm, ydesigner, data["Designer"])
    if data["Sub Number"] != "": 
        pdf_canvas.setFont("Helvetica-Bold", 11.0)
        pdf_canvas.drawRightString(xr - xm, ysubnum, data["Sub Number"])

def make_pdf(data, foutput):
    ''' Makes the pdf, with the given data.'''
    c = canvas.Canvas(foutput, letter)
    nx, ny = grid
    pm_top, pm_left, pm_bottom, pm_right = page_margin
    xbox, ybox = note_size  
    xbox_with_margin, ybox_with_margin = xbox + 2*note_margin, ybox + 2*note_margin 
    used_width, used_height = nx*xbox_with_margin + pm_left + pm_right, ny*ybox_with_margin + pm_top + pm_bottom
    xgap, ygap = (8.5 - used_width) / (nx - 1), (11.0 - used_height) / (ny - 1)
    irow = ny - 1
    icol = 0
    page_dirty = False
    npages = 0
    for d in data:
        loc = pm_left + icol * (xbox_with_margin + xgap), pm_bottom + irow * (ybox_with_margin + ygap)
        draw_note(d, loc, c)
        page_dirty = True
        icol += 1
        if icol >= nx:
            icol = 0 
            irow -= 1 
            if irow < 0:
                c.showPage()
                page_dirty = False 
                npages += 1 
                irow = ny - 1
    if page_dirty: 
        c.showPage()
        npages += 1
    try:
        c.save()
    except:
        print("ERROR: Unable to write pdf file %s.  (Is it opened in Adobe Reader?)" % foutput)
        return
    if npages <= 0:
        print("Weird Error... no pages output.")
    else:
        if npages == 1: print("PDF written (one page).") 
        else: print("PDF written to file %s (%d pages)." % (foutput, npages))  
        
def has_extension(fname, extensions):
    for e in extensions:
        if fname.endswith(e): return True
    
def run(finput, foutput):
    ''' Runs the program, with input and output filenames.'''
    if not has_extension(finput, (".xls", ".xlsx", ".XLS", ".XLSX")): finput += ".xlsx"
    if not has_extension(foutput, (".pdf", ".PDF")): foutput += ".pdf"
    if not os.path.isfile(finput):
        print("Unable to find file %s." % finput)
        return
    data = rddata(finput)
    if data is None: return
    if len(data) <= 0:
        print("No valid records (rows) found in the input file.")
        return
    make_pdf(data, foutput)


if __name__ == '__main__':
    if len(sys.argv) < 3 or len(sys.argv) > 4:
        print(Instructions)
        sys.exit()
    if len(sys.argv) == 4: mode = sys.argv[3]
    else: mode = "template"
    if mode == "template": do_outline = True 
    elif mode == "final":  do_outline = False
    else: 
        print('Improper mode.  Must be either "template" or "final".')
        print('(To see instructions, run program without arguments.)')
        sys.exit()
    finput = sys.argv[1]
    foutput = sys.argv[2]
    run(finput, foutput)
