from Tkinter import *
from tkFileDialog import askopenfilename
import os
import pickle
import xlrd
class App:
	def __init__(self,parent):

		f = Frame(parent)
		f.pack(padx=20,pady=10)
		
		self.text=Text(f,height=5)
		self.text.pack(side=TOP)

                self.output=Text(f,height=5)
		self.output.pack(side=BOTTOM)
		
		self.exit = Button(f, text="exit",command=f.quit)
		self.exit.pack(side=BOTTOM,padx=10,pady=10)

		self.upload=Button(f,text="upload",command=self.upload)
                self.upload.pack(side=BOTTOM,padx=10,pady=10)

                #self.entry = Entry(f,text="enter your choice")
		#self.entry.pack(side= TOP,padx=10,pady=12)
		
		self.button = Button(f, text="submit",command=self.print_this)
		self.button.pack(side=BOTTOM,padx=10,pady=10)
		
	def print_this(self):
		print "this is to be printed"
		contents=self.text.get(1.0, END)
		print contents
		contents=contents.split("\n")
		aspectcount=pickle.load(open('aspectcount.pkl','r'))
                for i in range(len(contents)):
                        if contents[i]=='':
                                continue
                        os.system("del customerreview.txt")
                        r=open("customerreview.txt",'a')
                        r.write(contents[i])
                        r.close()
                        os.system("python test.py")
                        book = xlrd.open_workbook("Training datatemp.xls")
                        sh = book.sheet_by_name("crsheet")
                        tot=book.sheet_by_name("total")
                        total=int(tot.cell_value(0,0))
                        ans=0
                        aspectsfound=set()
                        countdict=dict()
                        for r in range(len(aspectcount)):
                                countdict[r]=0
                        for r in range(sh.nrows):
                                aspectsfound.add(int(sh.cell_value(r,0)))
                                countdict[int(sh.cell_value(r,0))]+=int(sh.cell_value(r,2))
                        print countdict
                        print aspectsfound
                        for a in aspectsfound:
                                if countdict[a]>1:
                                        countdict[a]=1
                                if countdict[a]<-1:
                                        countdict[a]=-1
                                ans=ans+(countdict[a]*aspectcount[a])
                        ans=ans*1.0
                        ans=ans/total
                        self.output.insert(END,ans)
                        self.output.insert(END,"\n")

        def upload(self):
                filename = askopenfilename(filetypes=[("allfiles","*"),("pythonfiles","*.py")])
                print filename
                r=open(filename,'r')
                data=r.read()
                print data
                self.text.insert(END,data)
   		
root = Tk()
root.title('Sentiment Analysis')
Label(root,text='Enter the mobile review').pack(side=TOP,padx=10,pady=10)
app = App(root)

root.mainloop()
