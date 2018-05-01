import os
import pickle
import pandas
import urllib
import re
from nltk.tokenize import sent_tokenize,word_tokenize
from nltk.corpus import stopwords
import xlrd
import xlwt

masterD=pandas.DataFrame.from_csv('master_dict.csv', index_col=0)
uncerD=xlrd.open_workbook('uncertainty_dictionary.xlsx').sheet_by_index(0)
uncernL=[]
consL=[]
urlSt='https://www.sec.gov/Archives/'
rxl=xlrd.open_workbook('cik_list.xlsx')
rsheet=rxl.sheet_by_index(0)
osheet=xlwt.Workbook()
sheet=osheet.add_sheet('Output')
cols=['CIK','CONAME','FYRMO','FDATE','FORM','SECFNAME','mda_positive_score','mda_negative_score','mda_polarity_score','mda_average_sentence_length','mda_percentage_of_complex_words','mda_fog_index','mda_complex_word_count','mda_word_count','mda_uncertainty_score','mda_constraining_score','mda_positive_word_proportion',	'mda_negative_word_proportion',	'mda_uncertainty_word_proportion',	'mda_constraining_word_proportion',	'qqdmr_positive_score',	'qqdmr_negative_score',	'qqdmr_polarity_score', 'qqdmr_average_sentence_length',	'qqdmr_percentage_of_complex_words',	'qqdmr_fog_index',	'qqdmr_complex_word_count',	'qqdmr_word_count',	'qqdmr_uncertainty_score',	'qqdmr_constraining_score',	'qqdmr_positive_word_proportion',	'qqdmr_negative_word_proportion',	'qqdmr_uncertainty_word_proportion',	'qqdmr_constraining_word_proportion',	'rf_positive_score',	'rf_negative_score',	'rf_polarity_score',	'rf_average_sentence_length',	'rf_percentage_of_complex_words',	'rf_fog_index',	'rf_complex_word_count',	'rf_word_count',	'rf_uncertainty_score',	'rf_constraining_score',	'rf_positive_word_proportion',	'rf_negative_word_proportion',	'rf_uncertainty_word_proportion',	'rf_constraining_word_proportion',	'constraining_words_whole_report'
]
i=0
for col in cols:
    sheet.write(0,i,col)
    i=i+1
swords= stopwords.words('english')
for a in range(1,uncerD.nrows):
    uncernL.append(uncerD.cell(a,0).value)
consD=xlrd.open_workbook('constraining_dictionary.xlsx').sheet_by_index(0)
for b in range(1,consD.nrows):
    consL.append(consD.cell(b,0).value)

def write_to_osheet(row,col,value):
    sheet.write(row,col,value)

def constraining_whole_report():
    data=open(str(os.getcwd())+'\\temp\\temp.txt').readlines()
    data=word_tokenize(str(data))
    const=0
    for word in data:
        if word in consL:
            const=const+1
    return const

def section_extract(spat, epat):
    result=''
    rec=False
    for line in open(str(os.getcwd())+'\\temp\\temp.txt').readlines():
        if rec is False:
            if re.search(spat, str(line)) is not None:
                result=str(line)
                rec=True
        elif rec is True:
            if re.search(epat, str(line)) is not None:
                rec=False
            else:
                result= result+' '+ str(line)
    return result

def pop_swords(categ):
    nswords=''
    for w in word_tokenize(categ):
        if w not in swords:
            nswords=nswords+ " "+ w
    return nswords

#Pass dataset with no stopwords
def section_analysis(categ):
    senti_score={'pos':0,
    'neg':0,
    'uncern':0,
    'const':0
    }
    for word in word_tokenize(categ):
        if len(word)>2:
            try:
                if masterD.loc[word.upper()]['Positive']>0:
                    senti_score['pos']=senti_score['pos']+1
                if masterD.loc[word.upper()]['Negative']>0:
                    senti_score['neg']=senti_score['neg']+1
                if word.upper in uncernL:
                    senti_score['uncern']=senti_score['uncern']+1
                if word.upper() in consL:
                    senti_score['const']=senti_score['const']+1
            except:
                pass
    # print(senti_score)
    return senti_score

def AsentLength(dataset):
    nW=len(word_tokenize(dataset))
    nSents=len(sent_tokenize(dataset))
    # print(nW/nSents)
    return nW/nSents

#dataset without stopwords
def wordCount(dataset):
    noP=re.findall('[^!.? ]+',dataset)
    # print(len(noP))
    return len(noP)

def complexWords(dataset):
    nCW=0
    for word in word_tokenize(dataset):
        try:
            if masterD.loc[word]['Syllables']>=2:
                nCW= nCW + 1
        except:
            pass
    # print(nCW)
    return nCW

#% of complex words
def fogindex(sL,cW):
    # print(0.4*(sL+cW))
    return 0.4*(sL+cW)

def polarity_score(sents):
    pos=sents['pos']
    neg=sents['neg']
    return (pos-neg)/((pos+neg)+0.000001)

def aggregate_data(row,col,dataset):
    if dataset is not None:
        nostops=pop_swords(dataset)
        sentscore=section_analysis(nostops)
        sentL=AsentLength(dataset)
        wc=wordCount(nostops)
        cwor=complexWords(dataset)
        fog=fogindex(sentL,(cwor*100)/wc)
        polarity=polarity_score(sentscore)
        write_to_osheet(row,col,sentscore['pos'])
        col=col+1
        write_to_osheet(row,col,sentscore['neg'])
        col=col+1
        write_to_osheet(row,col,polarity)
        col=col+1
        write_to_osheet(row,col,sentL)
        col=col+1
        write_to_osheet(row,col,(cwor*100)/wc)
        col=col+1
        write_to_osheet(row,col,fog)
        col=col+1
        write_to_osheet(row,col,cwor)
        col=col+1
        write_to_osheet(row,col,wc)
        col=col+1
        write_to_osheet(row,col,sentscore['uncern'])
        col=col+1
        write_to_osheet(row,col,sentscore['const'])
        col=col+1
        write_to_osheet(row,col,sentscore['pos']/wc)
        col=col+1
        write_to_osheet(row,col,sentscore['neg']/wc)
        col=col+1
        write_to_osheet(row,col,sentscore['uncern']/wc)
        col=col+1
        write_to_osheet(row,col,sentscore['const']/wc)
        col=col+1
    else:
        cole=col+14
        for i in range(col,cole):
            write_to_osheet(row,i,'Null')


for i in range(1,rsheet.nrows):
    print(i)
    for x in range(rsheet.ncols):
        write_to_osheet(i,x,rsheet.cell(i,x).value)
    with open(str(os.getcwd())+'\\temp\\temp.txt','wb') as f:
        data=urllib.request.urlopen(urlSt+str(rsheet.cell(i,5).value)).read()
        f.write(data)

    cons=constraining_whole_report()
    write_to_osheet(i,48,cons)

    mda=section_extract("MANAGEMENT'S DISCUSSION AND ANALYSIS","ITEM [1-20].")
    qqmdr=section_extract("Quantitative and Qualitative Disclosures about Market Risk".upper(),"ITEM [1-20].")
    rf=section_extract("Risk Factors".upper(),"ITEM [1-20].")
    # print(mda)
    if len(mda)>1:
        aggregate_data(i,6,mda)
    else:
        aggregate_data(i,6,None)
    if len(qqmdr)>1:
        aggregate_data(i,20,qqmdr)
    else:
        aggregate_data(i,20,None)
    if len(rf)>1:
        aggregate_data(i,34,rf)
    else:
        aggregate_data(i,34,None)

osheet.save('Output1.xls')
