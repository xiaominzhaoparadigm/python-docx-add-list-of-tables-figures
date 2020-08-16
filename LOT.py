# -*- coding: utf-8 -*-
"""
Created on Sat Aug 15 21:40:26 2020

@author: Xiaomin Zhao
"""
from docx.oxml import OxmlElement as OE
from docx.oxml.ns import qn

def add_list_of_table(run):
    fldChar = OE('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    fldChar.set(qn('w:dirty'), 'true')
    instrText = OE('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\h \\z \\c "Table"'  #"Table" of list of table and "Figure" for list of figure
    fldChar2 = OE('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OE('w:t')
    fldChar3.text = "Right-click to update field."
    fldChar2.append(fldChar3)
    
    fldChar4 = OE('w:fldChar')
    fldChar4.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar4)
    