#!/usr/bin/env python
# 10/27/2015
# By Mark Henes
# Convert all .xls that are located in the current script directory and convert to .xlsx
import os
# pyexcel module imports
import pyexcel as pe
import pyexcel.ext.xls # import it to handle xls file
import pyexcel.ext.xlsx # import it to handle xlsx file

# Script path to the directory it is located.
pcwd=os.path.dirname(os.path.abspath(__file__))

# List files in directory by file extension
# Specify directory
items = os.listdir(pcwd)

# Specify extension in "if" loop
newlist = []
for names in items:
    if names.endswith(".xls"):
        newlist.append(names)

# Convert .xls to .xlsx
for i in range(len(newlist)):
    pyexcel.save_as(file_name=newlist[i], dest_file_name=newlist[i] + "x")
