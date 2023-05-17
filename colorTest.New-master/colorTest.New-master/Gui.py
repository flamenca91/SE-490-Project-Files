# from debug import debug
import logging
# import pdb
import docx
from docx import Document
from docx.shared import RGBColor
from docx.shared import Inches
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from typing import Tuple

from tkinter import scrolledtext
import re
import copy
import time

# This libraries are for opening word document automatically
import os
import platform
import subprocess

# This library is for opening excel document automatically
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt

import Targest2

def GUI1():
    try:
        # pdb.set_trace()
        # Creates a word document, saves it as "report 3, and also adds a heading
        
        # Creates the gui
        window = Tk(className=' TARGEST v.1.5.x ')
        # set window size #
        window.geometry("500x500")
        window['background'] = '#afeae6'

        # Creates button 1
        Button(window, text="Choose Document ", command=Targest2.generateReport, width = 26).pack()
        # Creates button 2
        global genRep
        genRep = Button(window, text="Generate Reports ", state= DISABLED, command=Targest2.generateReport2, width = 26)
        genRep.pack()
        global allTagsButton
        allTagsButton = Button(text="Open all tags Report", state= DISABLED, command=Targest2.getDocumentTable, width = 26)
        allTagsButton.pack()
        # Creates button 3
        global getDoc
        getDoc = Button(window, text="Open Generated Report", state= DISABLED, command=Targest2.getDocument, width = 26)
        getDoc.pack()
        # Creates Excel button button 4
        global getExcel
        getExcel = Button(text="Open Generated Excel Report", state= DISABLED, command=Targest2.createExcel, width = 26)
        getExcel.pack()
        # Creates button 5
        global getOrphan
        getOrphan = Button(text="Generate Orphan Report", state= DISABLED, command=Targest2.orphanGenReport, width = 26)
        getOrphan.pack()
        # Creates button 6
        global getOrphanDoc
        getOrphanDoc = Button(text="Open Orphan Tags Report", state= DISABLED, command=Targest2.getOrphanDocument, width = 26)
        getOrphanDoc.pack()
        
        # Creates button 7
        global button
        button = Button(text="End Program", command=window.destroy, width = 26)
        button.pack()

        # Create text widget and specify size.
        global Txt
        Txt = Text(window, height = 25, width = 55)
        Txt.pack()

        msg3 = ('You need a text file with paths to your documents\n 1. Please choose your documents by clicking on \n    the "choose document" button.\n 2. Click "Generate Reports".  \n\n')
        Txt.insert(tk.END, msg3) #print in GUI
        
    except Exception as e:
        # Log an error message
        logging.exception('main(): ERROR', exc_info=True)
    else:
        # Log a success message
        logging.info('main(): PASS')

        window.mainloop()
