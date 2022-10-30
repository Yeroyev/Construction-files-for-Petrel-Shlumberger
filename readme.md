# Python script for Petrel Schlumberger
Script generates 5 files (Casing.tub, Packer.tub, Perforation.ev, Plug.ev, Stimulation.ev, Tubing.tub) from the excel file to be imported into Petrel Shlumberger.

# Decription
Well_construction.py - script for proccessing excel file.
Example.xls - excel file with the example inside how to fill tables.
Output folder - Folder that is created when you run the script and includes all construction files.
    Casing_out.tub - contains 5 colums:
        Well ID - Name of the well.
        Date - the date when casing was made.
        Size - diametr of casing in mm.
        Intervals from - top interval of casing.
        Intervals to - base interval of casing.
    Packer_out.tub - contains 3 colums:
        Well ID - Name of the well.
        Date - the date when packer was made.
        Intervals from - top interval of packer.
    Perforation_out.ev - contains 4 colums:
        Well ID - Name of the well.
        Date - the date when perforation was made.
        Intervals from - top interval of perforation.
        Intervals to - base interval of perforation.
    Plug_out.ev - contains 3 colums:
        Well ID - Name of the well.
        Date - the date when plug was made.
        Intervals from - top interval of plug (fill plug from this interval to the bottom of the well).
    Stimulation_out.ev - contains 4 colums:
        Well ID - Name of the well.
        Date - the date when stimulating was made.
        Intervals from - top interval of stimulation.
        Intervals to - base interval of stimulation.
    Tubing_out.tub - contains 5 colums:
        Well ID - Name of the well.
        Date - the date when tubing was made.
        Size - diametr of tubing in mm.
        Intervals from - top interval of tubing.
        Intervals to - base interval of tubing.

# How to use
1. Copy 'Example.xls' and 'Well_construction.py' to your computer.
2. Fill 'Example.xls' with your own data according to the example in the file.
3. Run the script 'Well_construction.py'.
4. Folder 'Output' will be created with all files that you have filled in it.