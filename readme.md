# Part Numbers on Post-It Notes

This program is used to make post-it notes from a bill-of-materials (BOM).  

The post-it notes are used by our CNC and Machining guild to track
the parts being made.  They are usually made by the CAD team when parts
are submitted for manufacturing.

This program (written in python) reads a BOM in the form of an excel file,
and generates a PDF that can be printed with post-it notes afficxed to the
paper before printing.  This method causes the post-it notes to be printed
instead of hand-writen.

Better instructions can be obtained by running the program without providing
arguments.

To run the program, use a terminal window, and type:

    >python spn.py 

## Installing

The program can be installed into any scratch directory, usually "Scratch" under
documents.  Python must also be installed.  If it is not, use Google for
instructions on doing that.  Once python is installed, run the following 
commands:

    >pip install pandas
    >pip install openpyxl
    >pip install xlrd
    >pip install reportlab

These commands only need to be run once.