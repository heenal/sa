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
os.system("del \"Training datatemp.xls\"")
os.system("del *.out")
os.system("del *.test")
os.system("del *.train")
os.system("del *.model")

"""---open for reading---"""
book = xlrd.open_workbook("Training data8.xls")
sh = book.sheet_by_name("train")

"""---open for writing---"""
wbk=xlwt.Workbook()
sheet=wbk.add_sheet("datatrain")
wbk.save('Training datatemp1.xls')

book1 = xlrd.open_workbook("Training datatemp1.xls")
sh1 = book1.sheet_by_name("datatrain")

print sh.name, sh.nrows, sh.ncols

_spec_chars = [u'\xc1',u'\xe1',u'\xc0',u'\xc2',u'\xe0',u'\u043e',u'\xc2',u'\xe2',u'\xc4',u'\xe4',u'\xc3',u'\xe3',u'\xc5',u'\xe5',u'\xc6',u'\xe6',u'\xc7',u'\xe7',u'\xd0',u'\xf0',u'\xc9',u'\xe9',u'\xc8',u'\xe8',u'\xca',u'\xea',u'\xcb',u'\xeb',u'\xcd',u'\xed',u'\xcc',u'\xec',u'\xce',u'\xee',u'\xcf',u'\xef',u'\xd1',u'\xf1',u'\xd3',u'\xf3',u'\xd2',u'\xf2',u'\xd4',u'\xf4',u'\xd6',u'\xf6',u'\xd5',u'\xf5',u'\xd8',u'\xf8',u'\xdf',u'\xde',u'\xfe',u'\xda',u'\xfa',u'\xd9',u'\xf9',u'\xdb',u'\xfb',u'\xdc',u'\xfc',u'\xdd',u'\xfd',u'\xff',u'\xa9',u'\xae',u'\u2122',u'\u20ac',u'\xa2',u'\xa3',u'\u2018',u'\u2019',u'\u201c',u'\u201d',u'\xab',u'\xbb',u'\u2014',u'\u2013',u'\xb0',u'\xb1',u'\xbc',u'\xbd',u'\xbe',u'\xd7',u'\xf7',u'\u03b1',u'\u03b2',u'\u221e']

def cleanspec(s, cleaned=_spec_chars):
    return ''.join([(c in cleaned and ' ' or c) for c in s])

aspects = set()

for rx in range(sh.nrows):
    aspects.add(sh.cell_value(rx,1))
"""print aspects"""

"""for a in aspects:
    print a"""

"""---this will assign a number to each distinct feature ---"""

aspectdict = dict()
i=0
for a in aspects:
    aspectdict[a]=i
    i+=1
print aspectdict

aspectcount=dict()
for kt,vt in aspectdict.iteritems():
    aspectcount[int(vt)]=0

revaspectdict=dict()
for kt,vt in aspectdict.iteritems():
    revaspectdict[int(vt)]=kt
print revaspectdict
 
adj='JJ'
adv='RB'
verb='VB'

"""---set: gives unique value i.e. if you put in set: a,a,b It will contain a,b---"""
"""---dict: like a map in java. here maps feature to a number---"""
 
features=set()
featuredict=dict()
frequency=dict()
frequencyperaspect=[dict() for x in range(len(aspectdict))]

for r in range(sh.nrows):
    aspectcount[aspectdict[sh.cell_value(r,1)]]=aspectcount[aspectdict[sh.cell_value(r,1)]]+1
    text=nltk.word_tokenize(sh.cell_value(r,2).lower())
    for t in text:
        if t[-1]=='.':
            t=t[:-1]
    """print text"""
    tmp=text
    text=nltk.pos_tag(text)
    """print text"""
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
       if text[l][0] not in stopwords.words('english') and not (text[l][0]=='not') and not(text[l][0]=="n't"):
            """print text[l][0]"""
            s=st.stem(text[l][0])
            string=string+" "+s
            features.add(s)
            if s in frequency.keys():
                frequency[s]+=1
            else:
                frequency[s]=1
            if s in frequencyperaspect[aspectdict[sh.cell_value(r,1)]].keys():
                frequencyperaspect[aspectdict[sh.cell_value(r,1)]][s]+=1
            else:
                frequencyperaspect[aspectdict[sh.cell_value(r,1)]][s]=1
    sheet.write(r,0,string)


"""print features"""

wbk.save('Training datatemp1.xls')  

book1 = xlrd.open_workbook("Training datatemp1.xls")
sh1 = book1.sheet_by_name("datatrain")

print len(aspectdict)

"""---featureset is like a 2d matrix. 1row for each aspect. Each row contains columns equals to features used for that aspect.---"""

featureset=[set() for x in range(len(aspectdict))]

keepset=[set() for x in range(len(aspectdict))]
for r in  range(sh.nrows):
    asp=sh.cell_value(r,1)
    string=sh1.cell_value(r,0)
    string=string[1:]
    if len(string)>0:
        string=string.split(" ")
        """print string"""
        for i in range(len(string)):
            featureset[aspectdict[asp]].add(string[i])
            if string[i]=='graph':
                print asp
        """print featureset[aspectdict[asp]]"""

book = xlrd.open_workbook("Training data8.xls")
keep = book.sheet_by_name("keep")
for r in range(keep.nrows):
    asp=keep.cell_value(r,0)
    string=keep.cell_value(r,1)
    string=string.split(",")
    for s in string:
        keepset[aspectdict[asp]].add(st.stem(s))

"""print featuredict"""

"""for r in range(len(featureset)):
    print len(featureset[r])"""

aspectlexicon=[set() for x in range(len(aspectdict)*2)]                       
for r in range (sh.nrows):
    pol=sh.cell_value(r,0)
    asp=sh.cell_value(r,1)
    string=sh1.cell_value(r,0)
    string=string[1:]
    if len(string)>0:
        string=string.split(" ")
        if pol==1:
            for i in range(len(string)):
                aspectlexicon[aspectdict[asp]*2].add(string[i])
        else:
            for i in range(len(string)):
                aspectlexicon[aspectdict[asp]*2+1].add(string[i])
  
toremove=set()
for r in range(len(aspectdict)):
    for asp in featureset[r]:
        if frequency[asp]<4 or frequency[asp] >30:
            flag=0
            for rt in range(r+1, len(aspectdict)):
                if asp in featureset[rt]:
                    toremove.add(asp)


for asp in toremove:
    for r in range(len(aspectdict)):
        if asp in featureset[r]:
            if asp not in keepset[r]:
                featureset[r].remove(asp)
        """if asp in aspectlexicon[2*r]:
            if asp not in keepset[r]:
                aspectlexicon[2*r].remove(asp)
        if asp in aspectlexicon[2*r+1]:
            if asp not in keepset[r]:
                aspectlexicon[2*r+1].remove(asp)"""


""" removing lexicons common in +ve and -ve list of aspect """

"""for r in range(len(aspectdict)):
    removelexicon=set()
    for asp in aspectlexicon[2*r]:
        if asp in aspectlexicon[2*r+1]:
            removelexicon.add(asp)
    for asp in removelexicon:
        aspectlexicon[2*r].remove(asp)
        aspectlexicon[2*r+1].remove(asp)"""


i=1;
for r in range(len(aspectdict)):
    for asp in featureset[r]:
        featuredict[asp]=i
        i+=1


"""---open/create text file for appending ('a') ---"""
f=open('aspectdata.train','a')

"""---this part only writes to text file using aspectdict and featuredict formed above---"""

for r in range(sh.nrows):
    f.write(str(aspectdict[str(sh.cell_value(r,1))]))
    f.write(" ")
    temp=set()
    string=sh1.cell_value(r,0)
    string=string[1:]
    if len(string)>0:
        string=string.split(" ")
        for i in range(len(string)):
            temp.add(string[i])
    for key,value in sorted(featuredict.iteritems(), key=lambda (k,v): (v,k)):
        f.write(str(featuredict[key]))
        f.write(":")
        if key in temp:
            """f.write(str(frequency[key]))"""
            f.write("1 ")
        else:
            f.write("0 ")
    f.write("\n")  

"""create training files for each aspect (polarity detection of individual review) """

lexicondict=[dict() for x in range(len(aspectdict))] #combines aspectlexicon[2*i] and aspectlexicon[2*i+1] into lexicondict[i]

for kt,vt in aspectdict.iteritems():
    i=1
    for t in aspectlexicon[aspectdict[kt]*2]:
        lexicondict[vt][t]=i;
        i+=1;
    for t in aspectlexicon[aspectdict[kt]*2+1]:
        if t not in lexicondict[vt]:
            lexicondict[vt][t]=i;
            i+=1;
for r in range(sh.nrows):
    name="C:\\Users\\KH\\Desktop\\BE Project 16-3\\libsvm-3.16\\windows\\"+sh.cell_value(r,1)+".train"
    f=open(name, 'a')
    temp=set()
    string=sh1.cell_value(r,0)
    string=string[1:]
    if len(string)>0:
        f.write(str(sh.cell_value(r,0)))
        f.write(" ")
        string=string.split(" ")
        for i in range(len(string)):
            temp.add(string[i])
        for key,value in sorted(lexicondict[aspectdict[sh.cell_value(r,1)]].iteritems(), key=lambda (k,v): (v,k)):
            f.write(str(lexicondict[aspectdict[sh.cell_value(r,1)]][key]))
            f.write(":")
            if key in temp:
                f.write("1 ")
            else:
                f.write("0 ")
    f.write("\n")        
os.system("svm-train.exe -t 0 -q aspectdata.train")

os.system("del *.pkl")

output1 = open('aspectdict.pkl', 'wb')
pickle.dump(aspectdict, output1)
output1.close()
output1 = open('revaspectdict.pkl', 'wb')
pickle.dump(revaspectdict, output1)
output1.close()
output2 = open('lexicondict.pkl', 'wb')
pickle.dump(lexicondict, output2)
output2.close()

output1 = open('featuredict.pkl', 'wb')
pickle.dump(featuredict, output1)
output1.close()
output2 = open('aspectcount.pkl', 'wb')
pickle.dump(aspectcount, output2)
output2.close()

for kt,vt in aspectdict.iteritems():
    os.system("svm-train.exe -t 0 -q \""+kt+".train\"")


