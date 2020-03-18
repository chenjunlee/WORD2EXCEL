# -*- coding: utf-8 -*-
"""
Created on Sun Mar  8 15:53:02 2020

@author: Chenjun Li
"""

import docx
import os
import glob
import xlsxwriter

#################################
# Convert docx to a list of text
#################################
def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    question = True
    index = 0
    for para in doc.paragraphs:
        if '\u201c' in para.text: #Use the unicode of \"
            for run in para.runs:
                if run.font.size == 133350: #I don't know why this number, I just print out and find, it should be size of 12.5
                    question = False
            if question == False:
                fullText.insert(index, b"Answer")
                question = True
                index = index + 1
            else:
                fullText.insert(index, b"Question")
                index = index + 1
                
            txt = para.text.encode('ascii' , 'ignore')
            fullText.insert(index, txt)
            index = index + 1
        elif 'No text stored' in para.text:
            fullText.insert(index, b"Question")
            index = index + 1
            txt = para.text.encode('ascii' , 'ignore')
            fullText.insert(index, txt)
            index = index + 1
        elif 'Echo Show' in para.text:
            fullText.insert(index, b"Label")
            index = index + 1
            txt = para.text.encode('ascii' , 'ignore')
            fullText.insert(index, txt)
            index = 0
            
    return b'\n'.join(fullText)


def genTxt():
    filelist=glob.glob("*.docx")
    for f in filelist:
        name = os.path.basename(f)
        (filename, ext) = os.path.splitext(name)
        docname = filename + ext
        txtname = filename + ".txt"
        file = open(txtname, "w") 
        file.truncate(0)
        file.write(getText(docname).decode('utf-8')) 
        file.close()
        xlsname = filename + ".xlsx"
        workbook = xlsxwriter.Workbook(xlsname)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, 'User Words')
        worksheet.write(0, 7, 'Alexa Respond')
        worksheet.write(0, 20, 'Timestamp Label')
        lastQALabel = True
        row = 1
        with open(txtname) as fp:
            line = fp.readline()
            while line:
                if 'Question' in line:
                    line = fp.readline()
                    if lastQALabel:
                        worksheet.write(row, 0, line)
                        lastQALabel = False
                    else:
                        row = row + 1
                        worksheet.write(row, 0, line)
                        lastQALabel = False
                elif 'Answer' in line:
                    line = fp.readline()
                    worksheet.write(row, 7, line)
                elif 'Label' in line:
                    line = fp.readline()
                    worksheet.write(row, 20, line)
                    lastQALabel = True
                    row = row + 1
                line = fp.readline()
        workbook.close()
        os.remove(txtname)
        
genTxt()