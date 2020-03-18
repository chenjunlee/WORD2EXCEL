# WORD2EXCEL
Get needed info from word and transfer them to excel
Word is the Alexa voice record. The code generate txt file to filter picture. 
Get the required information and judge them by font size to know it is user
question or Alexa answer or timestamps. Label them in the txt file. Then read
txt file by label, import the data to Excel. Delete txt file.
With running it, put this file with the word file same directory, I use
spyder to run it. For some reason which I don't know, it doesn't work when
run it under commend line. 
Make sure you have: 
pip install python-docx
pip install xlsxwriter
