import xlrd
import nltk
import xlwt
from nltk.corpus import stopwords
from nltk.stem.lancaster import LancasterStemmer
st=LancasterStemmer()
"""---open for reading---"""
book = xlrd.open_workbook("Training data7.xls")
sh = book.sheet_by_name("train")

"""---open for writing---"""
wbk=xlwt.Workbook()
sheet=wbk.add_sheet("datatrain")
wbk.save('Training datatemp.xls')

book1 = xlrd.open_workbook("Training datatemp.xls")
sh1 = book1.sheet_by_name("datatrain")

print sh.name, sh.nrows, sh.ncols

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

"""aspects.add("camfera")
print aspects"""

adj='JJ'
adv='RB'
verb='VB'

def split1(txt, seps):
    default_sep = seps[0]
    for sep in seps[1:]: # we skip seps[0] because that's the default seperator
        txt = txt.replace(sep, default_sep)
    return [i.strip() for i in txt.split(default_sep)]


"""---set: gives unique value i.e. if you put in set: a,a,b It will contain a,b---"""
"""---dict: like a map in java. here maps feature to a number---"""
 
features=set()
featuredict=dict()
frequency=dict()

for r in range(sh.nrows):
    text=sh.cell_value(r,2)
    text=text.lower()
    sheet.write(r,2,text)
    text=nltk.word_tokenize(text)
    tmp=text
    t=' '.join(text)
    sheet.write(r,1,t)
    """print text"""
    text=nltk.pos_tag(text)
    #print text
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
       s=st.stem(text[l][0]) 
       if s not in stopwords.words('english') and not(text[l][0]=="n't") and not(text[l][0]=="not"):
            """print text[l][0]"""
            string=string+" "+s
            features.add(s)
            if s in frequency.keys():
                frequency[s]+=1
            else:
                frequency[s]=1
    sheet.write(r,0,string)


"""print features"""

wbk.save('Training datatemp.xls')  

book1 = xlrd.open_workbook("Training datatemp.xls")
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
        """print featureset[aspectdict[asp]]"""

book = xlrd.open_workbook("Training data7.xls")
keep = book.sheet_by_name("keep")
for r in range(keep.nrows):
    asp=keep.cell_value(r,0)
    string=keep.cell_value(r,1)
    string=string.split(",")
    for s in string:
        keepset[aspectdict[asp]].add(s)
        
for r in  range(sh.nrows):
    asp=sh.cell_value(r,1)
    string=sh1.cell_value(r,0)
    string=string[1:]
    if len(string)>0:
        string=string.split(" ")
        """print string"""
        for i in range(len(string)):
            featureset[aspectdict[asp]].add(string[i])
        """print featureset[aspectdict[asp]]"""


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


"""print featureset[4]
print aspectlexicon[8]
print aspectlexicon[9]
print len(aspectlexicon)"""
  
toremove=set()
for r in range(len(aspectdict)):
    for asp in featureset[r]:
        flag=0
        for rt in range(r+1, len(aspectdict)):
            if asp in featureset[rt]:
                toremove.add(asp)

   
 
"""print featureset[4]
print aspectlexicon[8]
print aspectlexicon[9]
print len(aspectlexicon)"""

""" removing lexicons common in +ve and -ve list of aspect """

for asp in toremove:
    for r in range(len(aspectdict)):
        if asp in featureset[r]:
            if asp not in keepset[r]:
                featureset[r].remove(asp)
            else:
                print asp
                print r
        if asp in aspectlexicon[2*r]:
            if asp not in keepset[r]:
                aspectlexicon[2*r].remove(asp)
        if asp in aspectlexicon[2*r+1]:
            if asp not in keepset[r]:
                aspectlexicon[2*r+1].remove(asp)   

"""print aspectlexicon[8]
print aspectlexicon[9]"""

"""for r in range(len(featureset)):
    print len(featureset[r])"""


i=1;
for r in range(len(aspectdict)):
    for asp in featureset[r]:
        featuredict[asp]=i
        i+=1

"""for i in frequency.keys():
    if frequency[i]>1 and i not in toremove:
        print i
        print featuredict[i]
"""
"""---open/create text file for appending ('a') ---"""
f=open('aspectdata.train','a')

"""f.write("abcdefg\n")
f.write("abefg")"""

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

lexicondict=[dict() for x in range(len(aspectdict))]
for kt,vt in aspectdict.iteritems():
    name="C:\\Users\\Fujitsu\\Desktop\\BE Project\\libsvm-3.16\\windows\\"+kt+".train"
    f=open(name, 'a')
    i=1
    for t in aspectlexicon[aspectdict[kt]*2]:
        lexicondict[vt][t]=i;
        i+=1;
    for t in aspectlexicon[aspectdict[kt]*2+1]:
        lexicondict[vt][t]=i;
        i+=1;
    for r in range(sh.nrows):
        if sh.cell_value(r,1)==kt:
            temp=set()
            string=sh1.cell_value(r,0)
            string=string[1:]
            if len(string)>0:
                f.write(str(sh.cell_value(r,0)))
                f.write(" ")
                string=string.split(" ")
                for i in range(len(string)):
                    temp.add(string[i])
                for key,value in sorted(lexicondict[vt].iteritems(), key=lambda (k,v): (v,k)):
                    f.write(str(lexicondict[vt][key]))
                    f.write(":")
                    if key in temp:
                        f.write("1 ")
                    else:
                        f.write("0 ")
            f.write("\n")

"""--- For testing ---"""

sh = book.sheet_by_name("test")

sheet=wbk.add_sheet("datatest")
wbk.save('Training datatemp.xls')

book1 = xlrd.open_workbook("Training datatemp.xls")
sh1 = book1.sheet_by_name("datatest")

features=set()

for r in range(sh.nrows):
    text=sh.cell_value(r,2)
    text=text.lower()
    sheet.write(r,2,text)
    text=nltk.word_tokenize(text)
    tmp=text
    t=' '.join(text)
    sheet.write(r,1,t)
    """print text"""
    text=nltk.pos_tag(text)
    if 'not' in tmp or "n't" in tmp:
        flag=0
        for l in range(len(text)):
            if "not"==text[l][0] or "n't"==text[l][0]:
                flag=l+1
                break
        for i in range(flag,len(text)):
            if text[i][1]==adv or text[i][1]==adj or text[i][1]==verb:
                tmp[i]="not_"+tmp[i]
                break
    text=nltk.pos_tag(tmp)
    """print text"""
    string=""
    for l in range(len(text)):
        if text[l][0] not in nltk.corpus.stopwords.words('english') and not(text[l][0]=="n't") and not(text[l][0]=="not"):
            """print text[l][0]"""
            string=string+" "+text[l][0]
            features.add(text[l][0])
    sheet.write(r,0,string)

"""print features"""
wbk.save('Training datatemp.xls')  

"""--- do not form any dict i.e.spectdict or featuredict from test ---"""

book1 = xlrd.open_workbook("Training datatemp.xls")
sh1 = book1.sheet_by_name("datatest")

featureset=[set() for x in range(len(aspectdict))]



for r in  range(sh.nrows):
    asp=sh.cell_value(r,1)
    string=sh1.cell_value(r,0)
    string=string[1:]
    if len(string)>0:
        string=string.split(" ")
        """print string"""
        for i in range(len(string)):
            featureset[aspectdict[asp]].add(string[i])
        

"""---open/create text file for appending ('a') ---"""
f=open('aspectdata.test','a')

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
            f.write("1 ")
        else:
            f.write("0 ")
    f.write("\n")  

for kt,vt in aspectdict.iteritems():
    name="C:\\Users\\Fujitsu\\Desktop\\BE Project\\libsvm-3.16\\windows\\"+kt+".test"
    f=open(name, 'a')
    for r in range(sh.nrows):
        if sh.cell_value(r,1)==kt:
            temp=set()
            string=sh1.cell_value(r,0)
            string=string[1:]
            if len(string)>0:
                f.write(str(sh.cell_value(r,0)))
                f.write(" ")
                string=string.split(" ")
                for i in range(len(string)):
                    temp.add(string[i])
                for key,value in sorted(lexicondict[vt].iteritems(), key=lambda (k,v): (v,k)):
                    f.write(str(lexicondict[vt][key]))
                    f.write(":")
                    if key in temp:
                        f.write("1 ")
                    else:
                        f.write("0 ")
            f.write("\n")


revaspectdict=dict()
for kt,vt in aspectdict.iteritems():
    revaspectdict[int(vt)]=kt
print revaspectdict
    
for r in range(sh.nrows):
    temp=set()
    string=sh1.cell_value(r,0)
    string=string[1:]
    name="C:\\Users\\Fujitsu\\Desktop\\BE Project\\libsvm-3.16\\windows\\"+sh.cell_value(r,1)+"1.test"
    f=open(name, 'a')
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
