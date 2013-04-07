import xlrd
import nltk
import xlwt
import os
import pickle
import shutil
import time
from nltk.corpus import stopwords
from nltk.stem.lancaster import LancasterStemmer
st=LancasterStemmer()
os.system("del *.out")
os.system("del *.test")

adv='RB'
adj='JJ'
verb='VB'

wbk=xlwt.Workbook()

"""--- For testing ---"""
revaspectdict=pickle.load(open('revaspectdict.pkl','r'))
aspectdict=pickle.load(open('aspectdict.pkl','r'))
aspectcount=pickle.load(open('aspectcount.pkl','r'))
featuredict=pickle.load(open('featuredict.pkl','r'))
lexicondict=pickle.load(open('lexicondict.pkl','r'))

sheet=wbk.add_sheet("datatest")
wbk.save('Training datatemp.xls')

rsheet=wbk.add_sheet("reviewtest")
crsheet=wbk.add_sheet("combinedreviewsheet")
wbk.save('Training datatemp.xls')

t=open("customerreview.txt",'r')
data=t.read().lower()
print data

print data[-1]
if data[-1]=='.':
    data=data[0:-1]

def split1(txt, seps):
    default_sep = seps[0]
    for sep in seps[1:]: # we skip seps[0] because that's the default seperator
        txt = txt.replace(sep, default_sep)
    return [i.strip() for i in txt.split(default_sep)]

data=split1(data,[', ','; ','. ',', but '])
print data


for i in range(len(data)):
    rsheet.write(i,0,data[i]) #reviewtest column 1 contains parts of reviews

wbk.save('Training datatemp.xls')  

book1 = xlrd.open_workbook("Training datatemp.xls")
sh1 = book1.sheet_by_name("datatest")

"""sh = book.sheet_by_name("test")"""
sh = book1.sheet_by_name("reviewtest")

features=set()

for r in range(sh.nrows):
    text=nltk.word_tokenize(sh.cell_value(r,0))
    for t in text:
        if t[-1]=='.':
            t=t[:-1]
    """print text"""
    tmp=text
    text=nltk.pos_tag(text)
    print text
    if 'not' in tmp or "n't" in tmp:
        flag=0
        for l in range(len(text)):
            if "not"==text[l][0] or "n't"==text[l][0]:
                flag=l+1
                break
        for i in range(flag,len(text)):
            if  text[i][1]==adj or text[i][1]==verb:
                tmp[i]="not_"+tmp[i]
                break
    text=nltk.pos_tag(tmp)
    string=""
    for l in range(len(text)):
        if text[l][0] not in nltk.corpus.stopwords.words('english') and not(text[l][0]=='not') and not(text[l][0]=="n't"):
            s=st.stem(text[l][0])
            """print text[l][0]"""
            string=string+" "+s
    sheet.write(r,1,string) # datatest column 2 contains tokenized features of parts of review for test 

"""print features"""
wbk.save('Training datatemp.xls')  

"""--- do not form any dict i.e. aspectdict or featuredict from test ---"""

book1 = xlrd.open_workbook("Training datatemp.xls")
sh1 = book1.sheet_by_name("datatest")
       

"""---open/create text file for appending ('a') ---"""
f=open('aspectdata.test','a')

"""---this part only writes to text file using aspectdict and featuredict formed above---"""
 
        
for r in range(sh.nrows):
    """f.write(str(aspectdict[str(sh.cell_value(r,1))]))"""
    f.write("0 ")
    temp=set()
    string=sh1.cell_value(r,1)
    print string
    string=string[1:]
    if len(string)>0:
        string=string.split(" ")
        for i in range(len(string)):
            temp.add(string[i])
    for key,value in sorted(featuredict.iteritems(), key=lambda (k,v): (v,k)):
        f.write(str(featuredict[key]))
        f.write(":")
        if key in temp:
            f.write("1 ")
        else:
            f.write("0 ")
    f.write("\n")  
f.close()

f=open("aspectdata.test",'r')
f.read()
f.close()

os.system("echo 'heenal'")

os.system("svm-predict.exe aspectdata.test aspectdata.train.model aspect.out")

"""---copy the output of aspect classification to datatest column 1---"""
f=open("aspect.out",'r')
data=f.read()
data=data.split("\n")

if data[len(data)-1]=='':
    data.pop()
print data


for r in range(sh.nrows):
    sheet.write(r,0,data[r]) #datatest column 1 contains aspect classification output


"""---combine the adjacent reviews if they are of same aspect and write combined parts in combinedreviewsheet---"""
s=sh.cell_value(0,0)
r=sh.nrows
i=0
w=0
for j in range(1,r):
    if data[i]==data[j]:
        s=s+" "+sh.cell_value(j,0)
    else:
        crsheet.write(w,1,s)
        crsheet.write(w,0,data[i])
        w=w+1
        i=j
        s=sh.cell_value(i,0)
crsheet.write(w,1,s)
crsheet.write(w,0,data[i]) #combinedreviewsheet contains aspect no. in column 1 and combined parts of review in column 2 
wbk.save('Training datatemp.xls')

sheet=wbk.add_sheet("crsheet")
wbk.save('Training datatemp.xls')
total=wbk.add_sheet("total")
wbk.save('Training datatemp.xls')

book1 = xlrd.open_workbook("Training datatemp.xls")
sh = book1.sheet_by_name("combinedreviewsheet")

for r in range(sh.nrows):
    text=nltk.word_tokenize(sh.cell_value(r,1))
    """print text"""
    tmp=text
    text=nltk.pos_tag(text)
    if 'not' in tmp or "n't" in tmp:
        flag=0
        for l in range(len(text)):
            if "not"==text[l][0] or "n't"==text[l][0]:
                flag=l+1
                break
        for i in range(flag,len(text)):
            if  text[i][1]==adj or text[i][1]==verb:
                tmp[i]="not_"+tmp[i]
                break
    text=nltk.pos_tag(tmp)
    """print text"""
    string=""
    for l in range(len(text)):
        if text[l][0] not in nltk.corpus.stopwords.words('english') and not(text[l][0]=='not') and not(text[l][0]=="n't"):
            """print text[l][0]"""
            s=st.stem(text[l][0])
            string=string+" "+s
    sheet.write(r,1,string)
   
    sheet.write(r,0,sh.cell_value(r,0)) #crsheet contains aspect no. in column 1 and tokenized combined parts in column 2
    

wbk.save('Training datatemp.xls')

book1 = xlrd.open_workbook("Training datatemp.xls")
sh1 = book1.sheet_by_name("crsheet")
print sh1.nrows



for r in range(sh1.nrows):
    name="C:\\Users\\KH\\Desktop\\BE Project 16-3\\libsvm-3.16\\windows\\"+revaspectdict[int(sh1.cell_value(r,0))]+".test"
    f=open(name, 'a')
    temp=set()
    string=sh1.cell_value(r,1)
    print string
    string=string[1:]
    if len(string)>0:
        f.write("0 ")
        string=string.split(" ")
        for i in range(len(string)):
            temp.add(string[i])
        for key,value in sorted(lexicondict[int(sh1.cell_value(r,0))].iteritems(), key=lambda (k,v): (v,k)):
            f.write(str(lexicondict[int(sh1.cell_value(r,0))][key]))
            f.write(":")
            if key in temp:
                f.write("1 ")
            else:
                f.write("0 ")
    f.write("\n")
    f.close()
  

l=['expensive','phone']
for a in l:
    for i in range(len(aspectdict)):
        if a in lexicondict[i]:
            print a+" "+revaspectdict[i]

checkset=set()
for r in range(sh1.nrows):
    checkset.add(int(sh1.cell_value(r,0)))


for i in checkset:
    os.system("svm-predict.exe \""+revaspectdict[i]+".test\" \""+revaspectdict[i]+".train.model\" \""+revaspectdict[i]+".out\"")

for i in checkset:
    f=open(revaspectdict[i]+".out",'r')
    data=f.read()
    data=data.split("\n")
    p=0;
    for r in range(sh1.nrows):
        if int(sh1.cell_value(r,0))==i:
            sheet.write(r,2,data[p])
            p=p+1;
            sheet.write(r,3,aspectcount[int(sh1.cell_value(r,0))])

wbk.save('Training datatemp.xls')
print aspectcount

tot=0
for kt,vt in aspectcount.iteritems():
    tot=tot+vt

total.write(0,0,tot)
wbk.save('Training datatemp.xls')

print tot

