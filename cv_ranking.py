# -*- coding: utf-8 -*-


##IMPORTING NECESSARY LIBRARIES
import requests
import io
import pdfminer
from io import BytesIO
from io import StringIO
import re
import pandas as pd
import numpy as np
import spacy
import pickle
import random
import os, math
from spacy.matcher import PhraseMatcher
from collections import Counter

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

##LOADING SPACY MODULE WITH PHRASE MATCHERS' LIBRARY
nlp = spacy.load('en_core_web_sm')
phrase_matcher = PhraseMatcher(nlp.vocab)

## READING EXCEL FILE
excel_data_df = pd.read_excel('input.xlsx', sheet_name=None)

# print whole sheet data
print(excel_data_df)
type(excel_data_df )



excell=list(excel_data_df.keys())
excell

##DATA REFINING, REPLACING EMPTY SPACES WITH NAN VALUE
print(excel_data_df['Sheet2'].replace(r'^\s*$', np.nan, regex=True))

## LOADING SHEET 2
excello=excel_data_df['Sheet2']
cvlink=excello['Links to CV'].tolist()
print(cvlink)

##LOADING SCRAPPED LI DATA IN A LIST WHERE LIST ELEMENTS ARE INDIVIDUAL DETAILS FOR A PERSON
lily=[]
linkedinparse=excel_data_df['Sheet4']
for lis in range (0, len(cvlink)):
    lily.append(linkedinparse.iloc[lis])
lily

## JOINING N DIMENTIONAL LIST, WITH ALL ITS ELEMENTS
sed=','
uplinkdata=[]
for up in range(0, len(cvlink)):
    try:
        try:
            uplinkdata.append(sed.join(lily[up]))
        except:
            exercet='exception '
            for sii in range (0, len(lily[up])):
                exercet=exercet+' '+str((lily[up][sii]))
            uplinkdata.append(exercet)
    except:
        uplinkdata.append(lily[up][1])

uplinkdata

## LOADING SHEET 3 FOR KEYWORDS
excel_datakey_df = excel_data_df['Sheet3']
col=excel_datakey_df.columns
# print whole sheet data
print(excel_datakey_df)

## COLUMN HEADINGS
col

## FOLLOWING FUNCTION CODE CONVERTS THE DATA 
##ON A CV INTO A LIST OF STRING, EACH CV DATA IS STORED IN REFINED VARIABLE
ola=[]
def convert_pdf_to_txt(fp):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = "utf-8"
    # codec ='ISO-8859-1'
    laparams = LAParams()
    device = TextConverter(
        rsrcmgr, retstr, laparams=laparams
    )

    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(
        fp,
        pagenos,
        maxpages=maxpages,
        password=password,
        caching=caching,
        check_extractable=True,
    ):
        interpreter.process_page(page)

    text = retstr.getvalue()
    device.close()
    retstr.close()
    text = str(text)
    text = text.replace("\\n", "")
    text = text.lower()
#     return text
    listing=[]
    listing.append(text)
    refined=[]
    for sub in listing:
        refined.append(sub.replace("\n", ""))
#     print('redifined of convert function : ', refined)
    keycount=[]

    
#     for k in range(0, len(col)):
    oops=parsing('health', newpat,refined)
    ola.append(oops)
    return ola

## THIS FUNCTION PARSES THE DATA FROM CV, OVERLAPS THE KEYWORDS IN MATCHER AND 
##CALCULATES THE VALUE. THE PATTERNS OBTAINED IN NEXT SNIPPET ARE THE MATCHING WORDS BEING SEARCHED IN A CV
ops=[]

def parsing(name, newpat,refined):
    phrase_matcher.add("health", newpat)
    
#     print('inside func ',refined[0])    
    sentence = nlp (refined[0])
    matched_phrases = phrase_matcher(sentence)
#     print(matched_phrases)
    keycount=[]
    for match_id, start, end in matched_phrases:
        string_id = nlp.vocab.strings[match_id]  
        span = sentence[start:end]                   
#         print(string_id, span.text)
        keycount.append(span.text)
#         print(keycount)
    counters= Counter(keycount).items()
    keycount=[]

    object_list = []
    for key, value in counters:
        entry = [key, value]
        object_list.append(entry)
    print('For CV no. '+ str(i) +':')
    print(object_list)
    
    return(object_list)
    return ops

##THIS IS THE PIPELINING OF PHRASE MATCHER AND PATTERNS TO BE SEARCHED IN CV. 
##THIS SNIPPET FORMS A PATTERN OF KEYWORDS WHICH WE ARE LOOKING FOR IN CV
new=[]
newpatit=[]
for j in range (0, len(col)):
#         phrase_matcher = PhraseMatcher(nlp.vocab)
        linked1=excel_datakey_df[col[j]].tolist()
#         print(linked1)
        linked1=[x for x in linked1 if x == x]
        linked1=[x.lower() for x in linked1]
        
        new.append(linked1)
        
        patterns = [nlp(text) for text in linked1]
        newpatit.append(patterns)

new

newpatit

newpat = [ itemzx for elemzx in newpatit for itemzx in elemzx]

##THE LOGIC IS SAME AS ABOVE EXCEPT THAT THIS IS BEING APPLIED ON 
##LI DATA THAT WE ALREADY STORED IN THE UPLINKDATA LIST
opsli=[]
def linkparsing(name, newpat,lipo, uplinkdata):
    phrase_matcher.add("health", newpat)
    
        
    sentencelink = nlp (uplinkdata[lipo].lower())
    matched_phrasesli = phrase_matcher(sentencelink)
#         print(matched_phrases)
    keycountlink=[]
    for match_idlink, startlink, endlink in matched_phrasesli:
        string_idlink = nlp.vocab.strings[match_idlink]  
        spanlink = sentencelink[startlink:endlink]                   
#         print(string_id, span.text)
        keycountlink.append(spanlink.text)
    counterslink= Counter(keycountlink).items()
    object_listlink = []
    for keyli, valueli in counterslink:
        entryli = [keyli, valueli]
        object_listlink.append(entryli)
    print('For CV no. LINKED MATCHES '+ str(lipo) +':')
    print(object_listlink)
    
    return(object_listlink)
    return opsli

##THIS IS THE FUNCTION CALL FOR LI MATCHING KEYWORDS USING NLP- SPACY PHRASE MATCHERS
olalink=[]
for lipo in range (0, len(uplinkdata)):
    try:
#         for k in range(0, len(col)):
        oopslinkk=linkparsing('health', newpat,lipo, uplinkdata)
        olalink.append(oopslinkk)
    except:
#         for ui in range (0, len(col)):
        print('For CV no. '+ str(lipo) +':')
        print('nan')
        olalink.append('nan')
        continue

olalink

finlinkk=[]
# indlink=len(col)-1
# for nli in range (indlink,len(olalink), len(col)):
#     finlinkk.append(olalink[nli])
finlinkk=olalink

finlinkk

##CALCULATING THE TOTAL POINTS A LI PROFILE GOT
totalslink=[]
def scorecalcilink (new,finlinkk):
    vallinkk=[]
    for nu in range(0, len(new)):
        for p in range(0, len(finlinkk)):
            for h in range ( 0, len(finlinkk[p])):
                
                if set((finlinkk[p][h])) & set(new[nu]):
                    vallinkk.append(finlinkk[p][h][1])
                else:
                    vallinkk.append(0)
            print(vallinkk)
            totalli=0
            for ele in range(0, len(vallinkk)):
                totalli = totalli + vallinkk[ele]
            totalslink.append(totalli)
            print('Totals :',totalslink)
            vallinkk=[]
            totalli=0
    return totalslink

totalsslinkk=scorecalcilink(new,finlinkk)

print(totalsslinkk)

##SEGREGATING THE TOTAL INTO THE NUMBER OF DIMENSIONS THAT WE HAVE PASSED
def chunk_based_on_sizelink(lst, n):
    for x in range(0, len(lst), n):
        each_chunk = lst[x: n+x]

        if len(each_chunk) < n:
            each_chunk = each_chunk + [None for y in range(n-len(each_chunk))]
        yield each_chunk

seglinkk=list(chunk_based_on_sizelink(totalsslinkk, len(finlinkk)))

print(seglinkk)

##THIS IS THE FUNCTION CALL FOR NLP-SPACY PHRASE MATHERS AND PATTERNS FOR CV DATA
for i in range (0,len(cvlink)):
    try:
        response_pdf = requests.get(cvlink[i])
        pdf_stream = io.BytesIO(response_pdf.content)
        convert_pdf_to_txt(pdf_stream)
    except Exception as e:
#         for u in range (0, len(col)):
        print('For CV no. '+ str(i) +':')
        print(e)
        ola.append('nan')
        continue

ola

len(ola)

fin=[]
# ind=len(col)-1
# for n in range (ind,len(ola), len(col)):
#     fin.append(ola[n])
fin=ola

print(fin)

new

new[0]

len(new)

fin

totals=[]
def scorecalci (new,fin):
    val=[]
    for nu in range(0, len(new)):
        for p in range(0, len(fin)):
            for h in range ( 0, len(fin[p])):
                
                if set((fin[p][h])) & set(new[nu]):
                    val.append(fin[p][h][1])
                else:
                    val.append(0)
            print(val)
            total=0
            for ele in range(0, len(val)):
                total = total + val[ele]
            totals.append(total)
            print('Totals :',totals)
            val=[]
            total=0
    return totals

totalss=scorecalci(new,fin)

print(totalss)

len(fin)

##SEGREGATED CV TOTAL SCORE INTO NUMBER OF DIMENTIONS
def chunk_based_on_size(lst, n):
    for x in range(0, len(lst), n):
        each_chunk = lst[x: n+x]

        if len(each_chunk) < n:
            each_chunk = each_chunk + [None for y in range(n-len(each_chunk))]
        yield each_chunk

seg=list(chunk_based_on_size(totalss, len(fin)))

print(seg)

##ASSIGNING THE WEIGHT TO CV CONTENT AND LI CONTENT
percentweight = excel_data_df['Sheet1']
print(percentweight)
percentileweigh=percentweight.drop(['Type', 'SR. No.'],axis=1)
print(percentileweigh)
percentileweigh.set_index('Dimension', 
              inplace = True)
#         print(excex)
perpercv = percentileweigh.loc["CV weight"]
finpercv=perpercv.values.tolist()
print(finpercv[0])
perperli = percentileweigh.loc["LI weight"]
finperli=perperli.values.tolist()
print(finperli[0])

# dimen_checks=dimen_check
# dimen_name=dimen_check['Dimension'].tolist()
# print(dimen_name)

print('seg', seg)
print('seglinkk', seglinkk)

## REFINING: IF CV IS ABSENT GIVES 100 % TO LI AND VICE VERSA
tempolistt=[]
for mm in range (0, len(seglinkk[0])):
    fakeadd=0
    for nn in range (0, len(seglinkk)):
        fakeadd=fakeadd+seglinkk[nn][mm]
    if (fakeadd==0):
        for nn in range (0, len(seglinkk)):
            seglinkk[nn][mm]=seg[nn][mm]
print('seg', seg)
print('seglinkk', seglinkk)

for mm in range (0, len(seg[0])):
    fakeaddcv=0
    for nn in range (0, len(seg)):
        fakeaddcv=fakeaddcv+seg[nn][mm]
    if (fakeaddcv==0):
        for nn in range (0, len(seg)):
            seg[nn][mm]=seglinkk[nn][mm]
print('seg', seg)
print('seglinkk', seglinkk)


        
for o in range (0, len(seg)):
    for xop in range (0, len(seg[0])):
        tempolistt.append((finpercv[0]*seg[o][xop])+(finperli[0]*seglinkk[o][xop]))
    globals()['string%s' % o] = tempolistt
    tempolistt=[]
    print('string%s' % o)

print(string0)

print(string1)

print(string2)

print(string3)

print(string4)

print(string5)

# print(string6)

weigh = excel_data_df['Sheet1']
weigh

for g in range (0, len(seg)):
#     globals()['upstring%s' % g]=[]
    mxacc=max(seg[g])
    print('Max of String'+ str(g))
    print(mxacc)
    globals()['upstring%s' % g]=mxacc

upstring2

lung=[]
for op in range (0,len(seg)):
    lung.append(globals()['upstring%s' % op])

lung[6]

weightss=weigh['Weight'].tolist()

weightss

for h in range (0, len(seg)):
    
    globals()['supstring%s' % h]=[]

##EVERYTHING OBTAINED AS SCORE IS BEING MULTIPLIED BY ITS WEIGHT
# try:
for h in range (0, len(seg)):
    for t in range (0, len(globals()['string%s' % h])):
        try:
            globals()['supstring%s' % h].append((globals()['string%s' % h][t]*weightss[h])/lung[h]) 
        except:
            globals()['supstring%s' % h].append((globals()['string%s' % h][t]*0) )
                


#     for h in range (0, len(seg)):
#         for t in range (0, len(globals()['string%s' % h])):
#             globals()['supstring%s' % h].append((globals()['string%s' % h][t]*0) )

print(supstring0)

ciao=excel_data_df['Sheet1']
typo=ciao['Type'].tolist()
print(typo)

opt=typo.count('Optional')
opt

ko=typo.count('Knock out Software')
ko

adv=typo.count('Advantage Software')
adv

## CALCULATING TOTAL OF SCORE IN ALL DIMENS 
scor=[]
jinx=0
for jo in range(0, len(supstring0)):
    for im in range (0, opt):
        jinx=jinx+globals()['supstring%s' % im][jo]
    scor.append(jinx)
    jinx=0

scor

len(scor)

scar=scor

scar

ko_start=opt
ko_start

## VALIDATING THE SCORE USING OUR DEFINED KOs
for zup in range (0, len(supstring0)):
    try:
        scor[zup]=(scor[zup]* globals()['supstring%s' % ko_start][zup])/globals()['supstring%s' % ko_start][zup]
    except:
        scor[zup]=0

print(scor)

# print(supstring6)

# print(supstring7)
len(col)

## ADVANCING THE SCORE USING OUR ADVANCENMENTS
avgadv=[]
try:
    
    zoo=0
    for zop in range(0, len(supstring0)):
        for zing in range(opt+ko, len(col)):
            zoo=zoo+globals()['supstring%s' % zing][zop]
            print(zoo)
        avgadv.append(zoo/(len(col)-(opt+ko)))
        zoo=0
except:
    for zop in range(0, len(supstring0)):
        avgadv.append(0)

print(avgadv)

## CALCULATING FINAL SCORE BY ADDING ADVANCEMENTS

finalee=[]
try:
    for finale in range (0, len(supstring0)):
        finalee.append(scor[finale]+(scor[finale]*avgadv[finale]))
except:
    finalee=scor

keyley=['Keywords from CV', 'Keywords from LI']
type(keyley)

dimen_check = excel_data_df['Sheet1']
print(dimen_check)
dimen_checks=dimen_check
dimen_name=dimen_check['Dimension'].tolist()
print(dimen_name)

## VALIDATING THE SCORE USING MUST HAVE CONSTARINTS
if ('OR CA' and 'OR ICWA' in dimen_name):    
    try:
        excelite=excel_data_df['Sheet2']
        calink=excelite['CA Year'].tolist()
        print('CA',calink)
#         icwas=excel_data_df['Sheet2']
        icwali=excelite['ICWA Year'].tolist()
        print("ICWA",icwali)
        lolexcexxx=dimen_checks.drop(['Type', 'SR. No.'],axis=1)
        print(lolexcexxx)

        excex=excel_data_df['Sheet1']
        lolexcexxx.set_index('Dimension', 
              inplace = True)
        print(excex)
#         fixcex=excex.set_index('Dimension', 
#               inplace = True)
        resultsz = lolexcexxx.loc["OR ICWA"]
        finresicwa=resultsz.values.tolist()
        resultsz2 = lolexcexxx.loc["OR CA"]
        finresca=resultsz2.values.tolist()
        icwa_string=finresicwa[0]
        icwayear_list = icwa_string.split(' ')
        icwa_object = map(int, icwayear_list)
        icwa_of_integers = list(icwa_object)
#         print(icwa_of_integers[0])
        print('ICWA bounds', icwa_of_integers)
        ca_string=finresca[0]
        cayear_list = ca_string.split(' ')
        ca_object = map(int, cayear_list)
        ca_of_integers = list(ca_object)
        print('CA bounds', ca_of_integers)
        
        finaleee=[]
        for cu in range(0,len(finalee)):

                        if type(icwali[cu])==int:
                            if (icwa_of_integers[0] <=icwali[cu]<=icwa_of_integers[1]):
                                finaleee.append(finalee[cu])
                            else:
                                finaleee.append(0)
                        elif type(calink[cu])==int:
                                if (ca_of_integers[0]<=calink[cu]<= ca_of_integers[1]):
                                    finaleee.append(finalee[cu])
                                else:
                                    finaleee.append(0)
                        else :
                            finaleee.append(0)               
        print('finaleee',finaleee)
    except:
        pass
if ('AND CA' and 'AND ICWA' in dimen_name): 
#     try:
        excelite=excel_data_df['Sheet2']
        calink=excelite['CA Year'].tolist()
        print('CA',calink)
#         icwas=excel_data_df['Sheet2']
        icwali=excelite['ICWA Year'].tolist()
        print("ICWA",icwali)

        excex=excel_data_df['Sheet1']
        lolexcexxx=dimen_checks.drop(['Type', 'SR. No.'],axis=1)
        print(lolexcexxx)
        lolexcexxx.set_index('Dimension', 
              inplace = True)
        print(excex)
        resultsz = lolexcexxx.loc['AND ICWA']
        print(resultsz,'resultsz')
        finresicwa=resultsz.values.tolist()
        print(finresicwa,'finresicwa')
        resultsz2 = lolexcexxx.loc["AND CA"]
        finresca=resultsz2.values.tolist()
        icwa_string=finresicwa[0]
        icwayear_list = icwa_string.split(' ')
        icwa_object = map(int, icwayear_list)
        icwa_of_integers = list(icwa_object)
#         print(icwa_of_integers[0])
        print('ICWA bounds', icwa_of_integers)
        ca_string=finresca[0]
        cayear_list = ca_string.split(' ')
        ca_object = map(int, cayear_list)
        ca_of_integers = list(ca_object)
        print('CA bounds', ca_of_integers)
        
        finaleee=[]
        for cu in range(0,len(finalee)):

                        if type(icwali[cu])==int:
                            if type(calink[cu])==int:
                                
                                if ((icwa_of_integers[0] <=icwali[cu]<=icwa_of_integers[1]) and (ca_of_integers[0]<=calink[cu]<= ca_of_integers[1]) ):
                                    finaleee.append(finalee[cu])
                                else:
                                    finaleee.append(0)
                                    
                            else:
                                finaleee.append(0)
              
                        else :
                            finaleee.append(0)               
        print('finaleee',finaleee)
if ('CA' in dimen_name):
        excelite=excel_data_df['Sheet2']
        calink=excelite['CA Year'].tolist()
        print('CA',calink)
        excex=excel_data_df['Sheet1']
        lolexcexxx=dimen_checks.drop(['Type', 'SR. No.'],axis=1)
        print(lolexcexxx)
        
        lolexcexxx.set_index('Dimension', 
              inplace = True)
        print(excex)
        resultsz2 = lolexcexxx.loc["CA"]
        finresca=resultsz2.values.tolist()
        ca_string=finresca[0]
        cayear_list = ca_string.split(' ')
        ca_object = map(int, cayear_list)
        ca_of_integers = list(ca_object)
        print('CA bounds', ca_of_integers)
        
        finaleee=[]
        for cu in range(0,len(finalee)):
        
            if type(calink[cu])==int:
                if (ca_of_integers[0]<=calink[cu]<= ca_of_integers[1]):
                     finaleee.append(finalee[cu])
                else:
                    finaleee.append(0)
            else:
                finaleee.append(0)
                
        print('finaleee',finaleee)
if ('ICWA' in dimen_name):
        excelite=excel_data_df['Sheet2']
        icwali=excelite['ICWA Year'].tolist()
        print("ICWA",icwali)
        excex=excel_data_df['Sheet1']
        lolexcexxx=dimen_checks.drop(['Type', 'SR. No.'],axis=1)
        print(lolexcexxx)
        lolexcexxx.set_index('Dimension', 
              inplace = True)
        print(excex)
        resultsz = lolexcexxx.loc["ICWA"]
        finresicwa=resultsz.values.tolist()
        
        
        icwa_string=finresicwa[0]
        icwayear_list = icwa_string.split(' ')
        icwa_object = map(int, icwayear_list)
        icwa_of_integers = list(icwa_object)
#         print(icwa_of_integers[0])
        print('ICWA bounds', icwa_of_integers)
        
        finaleee=[]
        for cu in range(0,len(finalee)):
    
        
            if type(icwali[cu])==int:
                if(icwa_of_integers[0] <=icwali[cu]<=icwa_of_integers[1]):
                     finaleee.append(finalee[cu])
                else:
                    finaleee.append(0)
            else:
                finaleee.append(0)
        print('finaleee',finaleee)

weightss

weigh

## VALIDATING EXP AND CTC
gigigi = lolexcexxx.loc["Experience"]
print(gigigi)
gigigi[0]

ctcgigi = lolexcexxx.loc["CTC"]
print(ctcgigi)
ctcgigi[0]

a_string = gigigi[0]
a_list = a_string.split(',')
map_object = map(int, a_list)


list_of_integers = list(map_object)
print(list_of_integers)

lowboundexp=list_of_integers[0]
lowboundexp

boundexpdwn=list_of_integers[1]
boundexpdwn

boundexpup=list_of_integers[2]
boundexpup

upboundexp=list_of_integers[3]
upboundexp

ctc_string = ctcgigi[0]
ctc_list = ctc_string.split(',')
mapc_object = map(int, ctc_list)


ctc_of_integers = list(mapc_object)
print(ctc_of_integers)

lowerctc = ctc_of_integers[0]
dwnctc = ctc_of_integers[1]
upctc=ctc_of_integers[2]
upperctc=ctc_of_integers[3]
print(lowerctc)
print(dwnctc)
print(upctc)
print(upperctc)

try:
    advscoring=excel_data_df['Sheet2']
    ctc=advscoring['CTC'].tolist()
    print(ctc)
except:
    pass

len(ctc)

try:
    advscoring2=excel_data_df['Sheet2']
    expe=advscoring2['Experience'].tolist()
    print(expe)
except:
    pass

len(expe)

ctcfin=[]
for i in range(0,len(ctc)):
# for i in range(0,10):

    if (ctc[i]==0 or ctc[i]=='nan'):
        ctcfin.append(finaleee[i])
    if (lowerctc<=ctc[i]<=upperctc):
        if (dwnctc<=ctc[i]<=upctc):
            ctcfin.append(finaleee[i]+(finaleee[i]*2))
        else:
            ctcfin.append(finaleee[i]*1.5)
    else:
        ctcfin.append(0)

ctcfin

expexp=[]
for i in range(0,len(expe)):
# for i in range(0,10):

    if (lowboundexp<=expe[i]<=upboundexp):
        if (boundexpdwn<=expe[i]<=boundexpup):
            expexp.append(ctcfin[i]+(ctcfin[i]*2))
        else:
            expexp.append(ctcfin[i]*1.5)
    else:
        expexp.append(0)
#     if (lower<expe[i]<=upper):
#         expexp.append(ctcfin[i]*2)
#     if (lower-2<expe[i]<=upper+3)
#         expexp.append(ctcfin[i])

expexp

keyley

maxxy=max(expexp)
maxxy

## ALLOWING FINAL SCORES AFTER ALL THE MUST HAVE THINGS AE APPLIED
finalscores=[]
for i in range (0, len(expexp)):
    try:
        finalscores.append((expexp[i]*100)/maxxy)
    except:
        finalscores.append(0)

finalscores

# to store
fin

#### STORING IN OUTPUT EXCEL
with pd.ExcelWriter("output.xlsx", engine="openpyxl", mode="a") as writer:
    df00 = pd.DataFrame({'Keywordsss': [keyley[0]] })
    df00.to_excel(writer,sheet_name='Sheet5',startcol= 0, index=False, header= False)
    df001= pd.DataFrame({'KeywordsssLI': [keyley[1]] })
    df001.to_excel(writer,sheet_name='Sheet5',startcol= 1, index=False, header= False)
    for i in range(0,len(fin)):  
    #     df = pd.DataFrame([fin[i]]).T
        df = pd.DataFrame({'Keywordsss': [fin[i]]})
        zf= pd.DataFrame({'KeywordsssLI': [finlinkk[i]]})
#         print(df)
        print(fin[i])
        df.to_excel(writer, startrow= i+1, sheet_name='Sheet5', index=False, header= False)
        zf.to_excel(writer, startrow= i+1,startcol= 1, sheet_name='Sheet5', index=False, header= False)

    for zi in range (0, len(col)):
        df = pd.DataFrame({ str([col[zi]]) : globals()['string%s' % zi]
                            })
        df.to_excel(writer, index= False, startcol= zi+2, sheet_name='Sheet5')
        
        df2= pd.DataFrame({ str([col[zi]])+'Cumulative' : globals()['supstring%s' % zi]
                            })
        df2.to_excel(writer, index= False, startcol= zi+len(col)+2, sheet_name='Sheet5')
    df3= pd.DataFrame({ 'Pre-Total' : scar,
                   'Valid Total_KO' : scor,
                       'Advanced Score' : finalee,
                       'Valid Total_CA' : finaleee,
                       

                   'Improved Score Exp' : expexp,
                       'Final Score' :finalscores
    
                    })
    df3.to_excel(writer, index= False, startcol= zi+len(col)+3, sheet_name='Sheet5')

out_df = pd.read_excel('output.xlsx', sheet_name=None)
out_df

outcon_df = pd.read_excel('output.xlsx', sheet_name=['Sheet2', 'Sheet5'])
outcon_df

outconhead_df = pd.read_excel('output.xlsx', sheet_name='Sheet2')
outconhead_df.columns

outconhead2_df = pd.read_excel('output.xlsx', sheet_name='Sheet5')
outconhead2_df.columns

all_dffs=pd.concat([outconhead_df,outconhead2_df],axis= 1)
all_dffs



# all_dffs= pd.concat(outcon_df, axis= 1, ignore_index=True)
# print(all_dffs)
sorted_dffs=all_dffs.sort_values(by='Final Score',ascending=False, ignore_index=True)
sorted_dffs

with pd.ExcelWriter("output.xlsx", engine="openpyxl", mode="a") as writer: 
    sorted_dffs.to_excel(writer, index= True ,sheet_name='Sheet6', header= True)





