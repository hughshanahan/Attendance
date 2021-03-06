#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jan 27 15:12:06 2020
Extract clicker IDs from a HTML file created by Turning Point
input: filename of html
output: list of IDs
@author: upac004
"""


def parseTPHtml(filename):
    from bs4 import BeautifulSoup
    with open(filename) as fp:
        soup = BeautifulSoup(fp,features="html.parser")
    
    body = soup.body
    tags = body.find_all()
    IDs = []
    N=len(tags)
    for i in range(0,N-1):
        tagb = tags[i]
        tagtd = tags[i+1]
        if tagb.name == "b" and tagb.string == "Responding Device:":
            if tagtd.name == "td":
                if tagtd['style'] == "padding-right: 50px; nowrap":
                    IDs.append(tagtd.string)
    
    return(IDs)

    
"""
From a directory go into all the sub-directories of the form CSXXXX or IYXXXX of it, parse the HTML output 
(generated by the TP app) and create a CSV for each file in a separate directory. The filename is updated
with the module name so it can all be stored in the same directory. 

input: directory pathToRead, directory pathToWrite, moveOldFiles = False 
output: integer representing number of files processed
"""
def parseAllHtml(pathToRead, pathToWrite, moveOldFiles = False ):
    
    import os
    import re
    from pathlib import Path
    
    n = 0
    
    files = os.listdir(pathToRead)
    pathf = Path(pathToRead)
    
    for f in files:
        match = re.match(r'^(CS|IY)(\d\d\d\d)$',f)
        if match:
            thisDir = pathf / f
            os.chdir(thisDir) 
            if moveOldFiles:
                if not "PROCESSED" in os.listdir():
                    os.mkdir("PROCESSED")
            for h in os.listdir():
                matchHtml = re.search(r'html$',h)
                if matchHtml:
                    IDs = parseTPHtml(h)
                    c = re.sub("\.html",".csv",h)
                    o = f + " " + c
                    outputFile = Path(pathToWrite) / o
                    n += 1
                    with open(outputFile, "w") as fp:
                        for i in IDs:
                            fp.write(i+"\n")
                        fp.close()
                    if moveOldFiles:
                        os.rename(h,Path("Processed") / h)
    return(n)
     
                        

                    
                    
            
            
        
