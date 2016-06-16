#*****************
#Specify variables
#*****************

#Specify name of document to import in quotes.
excelname='CEA_CU_2016_v3.xlsx'

#Specify name to save document as in quotes.
wordname='CEA_CU_2016_v3.docx'

#Specify the default number of repeat groups.
defaultrc=11

#Specify title of survey if different from "form_title."
formtitle='CEA'

#Import docx program
import openpyxl
import docx
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.enum.style import WD_STYLE
from docx.shared import Pt
from docx.shared import Inches
import string


#
#from docx import Document
#document = Document('MSME Mexico Auditing Questionnaire.docx')
#document.save('MSME Mexico Auditing Questionnaire_Saved.docx')
#print document

#Import excel document
wb = openpyxl.load_workbook(excelname)
survey=wb.get_sheet_by_name('survey')
choices=wb.get_sheet_by_name('choices')
settings=wb.get_sheet_by_name('settings')

#Create word document
document = Document()

#Title and initial info
print settings
print settings['A2'].value
if formtitle!='':
    formtitle=settings['A2'].value
document.add_heading(formtitle, 0)
intro1=document.add_paragraph('Start time:__________                                End time:__________')
intro2=document.add_paragraph('Device ID:____________________               Subscriber ID:____________________')
intro3=document.add_paragraph('Sim ID (Serial):____________________    Device phone number:____________________')

#Define printing function for questions not in a repeat group.
def QuestionState(query, tip, sort, chart, colnum):
    global number
    global newcol
    number += 1
    if tableyesno==0:
        qp=document.add_paragraph(str(number)+'. '+query)
        qp.paragraph_format.space_after = Pt(0)
        qp.paragraph_format.space_before = Pt(12)
        if tip is not None:
            qph=document.add_paragraph('')
            qph.add_run(tip).italic = True
    else:
        if tip!=None:
            fullquestion=query+' ('+tip+')'
        else:
            fullquestion=query
        if sort.partition(' ')[0]=='select_one':
            chart.cell(0, colnum).text=str(number)+'. SELECT ONE ('+sort.partition(' ')[2]+'): '+fullquestion
        elif sort.partition(' ')[0]=='select_multiple':
            chart.cell(0, colnum).text=str(number)+'. SELECT MULTIPLE ('+sort.partition(' ')[2]+'): '+fullquestion
        else:
            chart.cell(0, colnum).text=str(number)+'. '+fullquestion
        
def OptionList(paths, category):
    #global optnum
    #optnum += 1
    #print optnum
    #if optnum>25:
    #    optlet=string.ascii_lowercase[(optnum+1)/26]+string.ascii_lowercase[(optnum%26)-1]
    #else:
    #    optlet=string.ascii_lowercase[optnum]
    for y in range(2, choices.max_row):
        list_name=choices['A'+str(y)].value
        name=choices['B'+str(y)].value
        label=choices['C'+str(y)].value
        if list_name==paths:
            if category=='select_multiple':
                op=document.add_paragraph('__ '+str(name)+' - '+str(label))
            else:
                op=document.add_paragraph('[] '+str(name)+' - '+str(label))
            op.paragraph_format.space_after=Pt(0)
            op.paragraph_format.left_indent=Inches(0.5)
        
    

#Generate variable for which begin repeats should trigger a table.
def TableTime(repeatgroup):
    tabletime=0
    for x in range(8, survey.max_row):
        if survey['A'+str(x)].value=='begin repeat' and survey['B'+str(x)].value==repeatgroup:
            tabletime=1
        if survey['A'+str(x)].value=='begin repeat' and survey['B'+str(x)].value!=repeatgroup:
            if tabletime==1:
                return 0
        if survey['A'+str(x)].value=='end repeat':
            return tabletime

#Generate function that scans a phrase and replaces references to earlier fields.
def ReplaceRefs(phrase, mode):
    tempphrase=phrase
    if '${' in tempphrase and '}' in tempphrase:
        referring=0
        ref=''
        for n in range(2, len(tempphrase)):
            if tempphrase[n]=='}':
                referring=0
                if ref in qnumbers and mode=='Q':
                    tempphrase=tempphrase.replace('${'+ref+'}', ' _________ (Answer to Q'+str(qnumbers[ref])+') ')
                ref=''
            if referring==1:
                ref=ref+tempphrase[n]
            if tempphrase[n-1]=='$' and tempphrase[n]=='{':
                referring=1
    return tempphrase     

#Row by Row
tableyesno=0
number=0
qnumbers={}
repeat=0
def Program(a, b):
    for x in range(a, b):
        global tableyesno
        global repeat
        print number
        type=survey['A'+str(x)].value
        programmed=type in ['text', 'integer', 'geopoint'] or type.partition(' ')[0] in ['select_one', 'select_multiple']
        qnumbers[survey['B'+str(x)].value]=number
        question=survey['C'+str(x)].value
        print question
        if question!=None and programmed:
            question=ReplaceRefs(question, 'Q')
        print question

        hint=survey['D'+str(x)].value

        if type=='begin repeat':
            repeatcount=survey['O'+str(x)].value
            if not isinstance(survey['O'+str(x)].value, int):
                repeatcount=defaultrc
            document.add_heading(question, 2)
            repeat=repeat+1
            check=repeat
            d=0
            for z in range(x, b):
                print survey['A'+str(z)].value
                print 'check is '+str(check)+' and repeat is '+str(repeat)
                if survey['A'+str(z)].value=='end repeat' and check==repeat:
                    d=z+1
                    break
                elif survey['A'+str(z)].value=='begin repeat' and check!=repeat:
                    check=check+1
                elif survey['A'+str(z)].value=='end repeat':
                    check=check-1
            tableyesno=TableTime(survey['B'+str(x)].value)
            if tableyesno==1:
                if repeatcount!=None:
                    rowcount=repeatcount
                else:
                    rowcount=defaultrc
                table=document.add_table(rows=rowcount, cols=0)
                table.style='TableGrid'
                table.autofit=False
                table.add_column(180000)
                table.cell(0, 0).text='#'
                for n in range(1, rowcount):
                    table.cell(n, 0).text=str(n)+'.'
                newcol=1
                typedict={}

        if type=='end repeat':
            repeat=repeat-1
            if tableyesno==1:
                table.autofit=True
                for opt in typedict:
                    options=opt
                    choicetype=typedict[options]
                    o=document.add_paragraph('')
                    o.add_run(options).underline = True
                    o.paragraph_format.space_before=Pt(12)
                    OptionList(options, choicetype)        
            tableyesno=0

        if tableyesno==0:
            if type=='begin group':
                document.add_heading(question, 1)
            if type=='text':
                QuestionState(question, hint, type, '', '')
                ans=document.add_paragraph('_________________________________________________________________________________________________________')
                ans.paragraph_format.space_before=Pt(6)
            if type=='integer':
                QuestionState(question, hint, type, '', '')
                ans=document.add_paragraph('__________________________')
                ans.paragraph_format.space_before=Pt(6)
            if type.partition(' ')[0]=='select_one':
                question=question+' (select one)'
                QuestionState(question, hint, type, '', '')
                choicetype=type.partition(' ')[0]
                options=type.partition(' ')[2]
                OptionList(options, choicetype)
            if type.partition(' ')[0]=='select_multiple':
                question=question+' (select multiple)'
                QuestionState(question, hint, type, '', '')
                choicetype=type.partition(' ')[0]
                options=type.partition(' ')[2]
                OptionList(options, choicetype)
            if type=='note' and '${' not in question:
                QuestionState(question, hint, type, '', '')
            if type=='geopoint':
                QuestionState(question, hint, type, '', '')
                ans=document.add_paragraph("Latitude:  __  __* __' __\"")
                ans=document.add_paragraph("Longitude: __  __* __' __\"")
                ans=document.add_paragraph("Altitude:  ______m")
                ans=document.add_paragraph("Accuracy:  ______m")
        else:
            if type=='begin group':
                table.add_column(914400)
                table.cell(0, newcol).text='BEGIN GROUP: '+question
            if type=='end group':
                table.add_column(914400)
                table.cell(0, newcol).text='END GROUP'
            if type=='text' or type=='note':
                table.add_column(914400)
                QuestionState(question, hint, type, table, newcol)
            if type=='integer':
                table.add_column(360000)
                QuestionState(question, hint, type, table, newcol)
            if type.partition(' ')[0]=='select_one':
                table.add_column(914400)
                QuestionState(question, hint, type, table, newcol)
                typedict[type.partition(' ')[2]]='select_one'
            if type.partition(' ')[0]=='select_multiple':
                table.add_column(914400)
                QuestionState(question, hint, type, table, newcol)
                typedict[type.partition(' ')[2]]='select_multiple'
            if type=='geopoint':
                table.add_column(1600000)
                QuestionState(question, hint, type, table, newcol)
                for n in range(1, rowcount):
                    table.cell(n, newcol).text="Latitude:  __  __* __' __\" \n Longitude: __  __* __' __\" \n Altitude:  ______m \n Accuracy:  ______m"
            if programmed:
                newcol=newcol+1
                
        if repeat==1:
            c=x
            for i in range(0, repeatcount):
                Program(c, d)

Program(8, survey.max_row)
            
document.save(wordname)
