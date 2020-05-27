import re
import sys
import os
import openpyxl as xl
import openpyxl.utils.exceptions as ex
from docx import Document
from docx.shared import Inches

nameNOM = re.compile(r'nameNOM')
nameINIT = re.compile(r'nameINIT')
nameTVOR = re.compile(r'nameTVOR')
idS = re.compile(r'idS')
idN = re.compile(r'idN')
RET = re.compile(r'RET')
EDUCATION = re.compile(r'EDUCATION')
ADDRESS = re.compile(r'ADDRESS')
SNILS = re.compile(r'SNILS')
INN = re.compile(r'INN')
ACC = re.compile(r'ACC')
BANK = re.compile(r'BANK')
BIK = re.compile(r'BIK')
KOR = re.compile(r'KOR')
BIN = re.compile(r'BIN')
KPP = re.compile(r'KPP')


def docx_replace_regex(doc_obj, regex, replace):

    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex, replace)


def fill_docs(data_filename, template_filename,path):
    try:
        data = xl.load_workbook(data_filename)
    except ex.InvalidFileException:
        print("Wrong file")
        sys.exit()
    except TypeError:
        sys.exit()
    ws = data.worksheets[0]
    doc = Document(template_filename)
    if not os.path.exists(path):
        os.mkdir(path)
    for r in ws.rows:
        if r[0].value:
            r = list(r)
            init = r[0].value.split()
            if len(init) == 3:
                docx_replace_regex(doc, nameINIT, r''+list(init[1])[0] + '. ' + list(init[2])[0] + '. ' + init[0])
            elif len(init)==2:
                docx_replace_regex(doc, nameINIT, r''+list(init[1])[0] + '. ' + init[0])
            else:
                docx_replace_regex(doc, nameINIT,r''+init[0])
            docx_replace_regex(doc, nameNOM, r''+str(r[0].value))
            docx_replace_regex(doc, nameTVOR, r''+str(r[1].value))
            docx_replace_regex(doc, idS, r''+str(r[2].value))
            docx_replace_regex(doc, idN, r''+str(r[3].value))
            docx_replace_regex(doc, RET, r''+str(r[4].value))
            docx_replace_regex(doc, EDUCATION, r''+str(r[5].value))
            docx_replace_regex(doc, ADDRESS, r''+str(r[6].value))
            docx_replace_regex(doc, SNILS, r''+str(r[7].value))
            docx_replace_regex(doc, INN, r''+str(r[8].value))
            docx_replace_regex(doc, ACC, r''+str(r[9].value))
            docx_replace_regex(doc, BANK, r''+str(r[10].value))
            docx_replace_regex(doc, BIK, r''+str(r[11].value))
            docx_replace_regex(doc, KOR, r''+str(r[12].value))
            docx_replace_regex(doc, BIN, r''+str(r[13].value))
            docx_replace_regex(doc, KPP, r''+str(r[14].value))
            doc.save(path + r[0].value + '.docx')
            doc = Document(template_filename)
