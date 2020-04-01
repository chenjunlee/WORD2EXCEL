# -*- coding: utf-8 -*-
"""
Created on Sun Mar 8 2020
Edited on Mar 31

@author: Chenjun Li
"""

import docx
import os
import glob

#################################
# Convert docx to a list of text
#################################
def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    question = True
    index = 0
    for para in doc.paragraphs:
        if '\u201c' in para.text: #Use the unicode of front double quotes
            for run in para.runs:
                if run.font.size == 133350: #I don't know why this number, I just print out and find, it should be size of 12.5 in word
                    question = False
            if question == False:
                fullText.insert(index, b"*VAS:\t")
                index = index + 1
                question = True
                txt = para.text.encode('ascii' , 'ignore')
                fullText.insert(index, txt)
                index = index + 1
            else:
                fullText.insert(index, b"*PAR:\t")
                index = index + 1
                txt = para.text.encode('ascii' , 'ignore')
                fullText.insert(index, txt)
                index = index + 1
                if "what" in para.text or "why" in para.text or "how" in para.text or "where" in para.text or "when" in para.text or "who" in para.text or "whose" in para.text or "which" in para.text:
                    fullText.insert(index, b" ?\n")
                    index = index + 1
                else:
                    fullText.insert(index, b" .\n")
                    index = index + 1
        elif 'Echo Show' in para.text:
            index = 0
            
    return b'\n'.join(fullText)

def endSymbol(x):
    end = len(x)-1
    lastCh = " ";
    res = x;
    while end >= 0:
        lastCh = x[end]
        if lastCh.isdigit() == False and lastCh.isalpha() == False and lastCh != " ":
            res = x[0:end] + " " + lastCh
            break
        elif lastCh == " ":
            end = end -1
        else:
            res = x[0:end+1]  + " ."
            break
    return res

def checkTime(x):
    if "a.m." in x:
        x = x.replace("a.m.", "a_m")
    if "p.m." in x:
        x = x.replace("p.m.", "p_m")
    return x

def replaceOK(x):
    if x == " ." or x == " ?":
        return x
    x = x.split(" ")
    result = "";
    start = True
    for word in x:
        if word == "ok":
            word = "okay"
        if word == "ok,":
            word = "okay,"
        if start == True:
            start = False
            result = result + word
        else:
            result = result + " " + word
    return result

ones = ["", "one ","two ","three ","four ", "five ", "six ","seven ","eight ","nine ",
        "ten ","eleven ","twelve ", "thirteen ", "fourteen ", "fifteen ","sixteen ","seventeen ", "eighteen ","nineteen "]
 
twenties = ["","","twenty ","thirty ","forty ", "fifty ","sixty ","seventy ","eighty ","ninety "]
 
thousands = ["","thousand ","million ", "billion ", "trillion ", "quadrillion ", "quintillion ", 
             "sextillion ", "septillion ","octillion ", "nonillion ", "decillion ", "undecillion ", 
             "duodecillion ", "tredecillion ", "quattuordecillion ", "quindecillion", "sexdecillion ", 
             "septendecillion ", "octodecillion ", "novemdecillion ", "vigintillion "]
 
def num999(n):
    c = n % 10 # singles digit
    b = ((n % 100) - c) / 10 # tens digit
    a = ((n % 1000) - (b * 10) - c) / 100 # hundreds digit
    t = ""
    h = ""
    if a != 0 and b == 0 and c == 0:
        t = ones[a] + "hundred "
    elif a != 0:
        t = ones[a] + "hundred and "
    if b <= 1:
        h = ones[n%100]
    elif b > 1:
        h = twenties[int(b)] + ones[int(c)]
    st = t + h
    return st
 
def num2word(num):
    if num == 0:
        return 'zero'
    i = 3
    n = str(num)
    word = ""
    k = 0
    while(i == 3):
        nw = n[-i:]
        n = n[:-i]
        if int(nw) == 0:
            word = num999(int(nw)) + thousands[int(nw)] + word
        else:
            word = num999(int(nw)) + thousands[k] + word
        if n == '':
            i = i+1
        k += 1
    return word[:-1]

def digitalToWords(x):
    start = True
    result = ""
    strs = x.split(" ")
    for s in strs:
        temp = ""
        if s.endswith(","):
            temp = ","
            s = s[0:len(s)-1]
        if s.isdigit() == True:
            s = num2word(int(s))
        if start == True:
            start = False
            result = result + s + temp
        else:
            result = result + " " + s + temp
    return result

def showTime(h, m): 
    nums = ["zero", "one", "two", "three", "four", 
            "five", "six", "seven", "eight", "nine", 
            "ten", "eleven", "twelve", "thirteen", 
            "fourteen", "fifteen", "sixteen",  
            "seventeen", "eighteen", "nineteen",  
            "twenty", "twenty one", "twenty two",  
            "twenty three", "twenty four",  
            "twenty five", "twenty six", "twenty seven", 
            "twenty eight", "twenty nine"]; 
    
    time = ""
    if m == 0: 
        time = nums[h] + " o'clock" 
    elif m == 1: 
        time = "one minute past " + nums[h] 
    elif m == 59: 
        time = "one minute to " + nums[(h % 12) + 1]
    elif m == 15: 
        time = "quarter past " + nums[h] 
    elif m == 30:
        time = "half past " + nums[h]
    elif m == 45: 
        time = "quarter to " + nums[(h % 12) + 1] 
    elif m <= 30: 
        time = nums[m] + " minutes past " + nums[h]
    elif m > 30: 
        time = nums[60 - m] + " minutes to " + nums[(h % 12) + 1]
    return time

def checkHour(x):
    x = x.split(" ")
    result = "";
    start = True
    temp1 = 0
    temp2 = 0
    temp = ""
    for s in x:
        if ":" in s:
            length = len(s)
            index = s.find(":")
            if index < length-1 and s[0:index].isdigit() == True and s[index+1:].isdigit() == True:
                temp1 = int(s[0:index])
                temp2 = int(s[index+1:])
                temp = showTime(temp1, temp2)
            else:
                temp = s
        else:
            temp = s
        if start == True:
            start = False
            result = result + temp
        else:
            result = result + " " + temp
    return result
                
                

def checkSpecial(x):
    x = x.replace("january", "January")
    x = x.replace("february", "February")
    x = x.replace("march", "March")
    x = x.replace("april", "April")
    x = x.replace("may", "May")
    x = x.replace("june", "June")
    x = x.replace("july", "July")
    x = x.replace("august", "August")
    x = x.replace("september", "September")
    x = x.replace("october", "October")
    x = x.replace("november", "November")
    x = x.replace("december", "December")
    x = x.replace("alexa", "Alexa")
    x = x.replace("thanksgiving", "Thanksgiving")
    x = x.replace("S. Y. M. P. T. O. M", "S Y M P O M")
    x = x.replace("reference.com", "reference dot com")
    if x.startswith("+"):
        temp = ""
        for c in x:
            if c == "+":
                temp = "Calling"
            else:
                temp = temp + " " + num2word(int(c))
        x = temp
    return x

def genTxt():
    filelist=glob.glob("*.docx")
    for f in filelist:
        name = os.path.basename(f)
        (filename, ext) = os.path.splitext(name)
        docname = filename + ext
        txtname = filename + ".txt"
        file = open(txtname, "w") 
        textList = ["@UTF8","@Begin","@Languages:\teng","@Participants:\tPAR Participant, VAS Media","@ID:\teng|change_me_later|PAR|||||Participant|||","@ID:\teng|change_me_later|VAS|||||Media|||"]
        for line in textList:
            file.write(line)
            file.write("\n")
        lines = getText(docname).decode('utf-8')
        rows = lines.split('\n')
        start = True
        substart = False
        par = True
        label = True
        for row in rows:
            if '*PAR' in row and start == False:
                label = True
                par = True
                file.write('\n')
            elif '*VAS' in row:
                label = True
                par = False
                substart = True
                file.write('\n')
            if label == True:
                label = False
                file.write(row)
                continue
            if par == False:
                row = checkSpecial(row)
                utterances = row.split('.')
                for utterance in utterances:
                    if utterance.startswith(" "):
                        utterance = utterance[1:]
                    utterance = utterance.lower();
                    utterance = digitalToWords(utterance)
                    utterance = endSymbol(utterance)
                    utterance = replaceOK(utterance)
                    utterance = checkSpecial(utterance)
                    utterance = checkHour(utterance)
                    if substart == True:
                        substart = False
                        file.write(utterance)
                    elif utterance != "":
                        file.write("\n")
                        file.write("*VAS:\t")
                        file.write(utterance)
            else:
                row = row.lower();
                row = checkTime(row);
                row = replaceOK(row);
                row = checkSpecial(row);
                file.write(row)
                
            start = False 

        file.write("\n@End\n")
        file.close()
        
genTxt()