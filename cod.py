#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct  4 12:24:25 2018

@author: nwit
"""
###Generates top and bottom oligos to clone guide sequences (minus PAM) into vector like pX458 (BbsI/BpiI overhangs)###
#Intructions: copy each guide sequence in one cell below eachother in one column in a file called input.csv
#Double-click script to run it directly in the terminal
#Output file (CRISPR_oligo_designer_output.xls) can be found in the same folder as the script


from Bio.Seq import Seq
#from Bio.Alphabet import generic_dna
import xlwt
import sys
from inspect import currentframe, getframeinfo
from pathlib import Path
import pandas as pd

#defines the location of the script (needed for running the script by double clicking on the script file):
filename = getframeinfo(currentframe()).filename
parent = Path(filename).resolve().parent
parent_str = str(parent)
file_location = parent_str + "/" + "input.csv"

#copies gene names and guide sequences from input.csv to a list:
csv = pd.read_csv(file_location)
gene_names = list(csv.gene) #gets data from comlumn with header "gene"
input_list = list(csv.guide_sequence) #gets data from comlumn with header "guide_sequence"

#checks if guide sequences are 20 nts long (stops script if one sequence is not 20 nts):
choice = input("Are input guide sequences truncated (<20 nts) (y/n + Enter)?:") #by answering with "y", this would allow for the design of oligos for truncated guide sequences
if choice == "n":
    for k in input_list:
        while True:
            if len(k) != 20:
                sys.exit("One or more guide sequence(s) are not 20 nts long, check input file")
            
            if len(k) == 20:
                break
else:
    pass

#generates top oligos in new list:
top_oligo_list = []
for c in input_list:
    d = c.upper()
    top = "CACCG" + d
    top_oligo_list.append(top)

#generates reverse complement oligos in new list:
rc_oligo_list = []
for e in input_list:
    r = e[::-1]
    dna = Seq(r)
    rc = dna.complement()
    str_rc = str(rc)
    rc_oligo_list.append(str_rc)

#adds overhangs to reverse complement oligos in new list:
bottom_oligo_list = []
for f in rc_oligo_list:
    bottom_oligo = "AAAC" + f + "C"
    bottom_oligo_list.append(bottom_oligo)

#creates oligo name lists:
gene_names_top = []
for aa in gene_names:
    aaa= aa + "_t"
    gene_names_top.append(aaa)

gene_names_bottom = []
for bb in gene_names:
    bbb= bb + "_b"
    gene_names_bottom.append(bbb)

#export oligo sequences to Excel sheet:
wb = xlwt.Workbook()
ws = wb.add_sheet('Oligo_output')
ws.write(0,0,"Oligo Name")
ws.write(0,1,"5'Mod")
ws.write(0,2,"Sequence")
ws.write(0,3,"3'Mod")
ws.write(0,4,"Scale (umole)")
ws.write(0,5,"Purification")
ws.write(0,6,"Format")
ws.write(0,7,"Concentration (uM)")
ws.write(0,8,"Number of tubes")
ws.write(0,9,"Notes")

i = 1
j = 2
for w,x,y,z in zip(gene_names_top,top_oligo_list,gene_names_bottom,bottom_oligo_list):
    ws.write(i,0,w)
    ws.write(i,2,x)
    ws.write(j,0,y)
    ws.write(j,2,z)
    i += 2
    j += 2
save_location = parent_str + "/" + "CRISPR_oligo_designer_output.xls"
wb.save(save_location)

print("Oligo design successful")
