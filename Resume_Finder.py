from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
import subprocess, os, platform
import docxpy
import sys
import shutil


''' Hyperlink helper class to handle click event on Text widget '''
class HyperlinkManager:
	def __init__(self, text):
		self.text = text
		self.text.tag_config("hyper", foreground="blue", underline=1)
		self.text.tag_bind("hyper", "<Enter>", self._enter)
		self.text.tag_bind("hyper", "<Leave>", self._leave)
		self.text.tag_bind("hyper", "<Button-1>", self._click)
		self.reset()
		
	def reset(self):
		self.links = {}
		
	def add(self, action):
		# add an action to the manager.  returns tags to use in
        # associated text widget
		tag = "hyper-%d" % len(self.links)
		self.links[tag] = action
		return "hyper", tag
		
	def _enter(self, event):
		self.text.config(cursor="hand2")
		
	def _leave(self, event):
		self.text.config(cursor="")
		
	def _click(self, event):
		for tag in self.text.tag_names(CURRENT):
			#print(tag)
			if tag[:6] == "hyper-":
				self.links[tag](tag)
				return


''' reset form '''
def reset(pathField,keywordsField, processingTextField, resultTextField):
	pathField.delete(0,END)
	keywordsField.delete(0,END)
	
	processingTextField.config(state=NORMAL)
	processingTextField.delete(1.0,END)
	processingTextField.insert('end', 'Re-initialized successfully..\n')
	processingTextField.config(state=DISABLED)
	
	resultTextField.config(state=NORMAL)
	resultTextField.delete(1.0,END)
	resultTextField.insert('end', 'Result will be dispay here\n')
	resultTextField.config(state=DISABLED)


''' browse directory '''
def browse(pathField, processingTextField, resultTextField):
	selected_directory = filedialog.askdirectory()
	pathField.delete(0, 'end')
	pathField.insert(0,selected_directory)	
	
	if(len(pathField.get()) != 0):
		processingTextField.config(state=NORMAL)
		processingTextField.insert(END,"Path Selected: '"+selected_directory+"'\n")
		processingTextField.config(state=DISABLED)


''' submit & process the files '''
def submit(pathField, keywordsField, processingTextField, resultTextField):
	
	directoryPath = pathField.get()
	if(len(directoryPath) == 0 or os.path.isdir(directoryPath) == False):
		messagebox.showwarning('File Error' , "Please select a directory")
		pathField.focus()
		
		processingTextField.config(state=NORMAL)
		processingTextField.insert(END,"Invalid Path Selected: '"+directoryPath+"'\n", 'error')
		processingTextField.config(state=DISABLED)
		return
	
	keywords = keywordsField.get()
	if(len(keywords) == 0):
		messagebox.showwarning('Invalid Argument Error',"Please provide at least a single keyword")
		keywordsField.focus()
		
		processingTextField.config(state=NORMAL)
		processingTextField.insert(END,"Invalid Argument Error: 'Please provide at least a single keyword'\n", 'error')
		processingTextField.config(state=DISABLED)
		return
	
	
	''' collect all variables that will used in further program ''' 
	searchOption = var.get()
	path = directoryPath + '/'
	keyword = keywords.lower()
		
		
	''' update processing messages.. '''
	processingTextField.config(state=NORMAL)
	processingTextField.insert(END,"\nFiles Processing..\n")
	processingTextField.config(state=DISABLED)
	
	list_files = os.listdir(path)
	
	processingTextField.config(state=NORMAL)
	processingTextField.insert(END,"Total ")
	processingTextField.insert(END,""+ str(len(list_files)), 'success')
	processingTextField.insert(END," files found..\n\n")
	processingTextField.config(state=DISABLED)
	print("Total no. of Files:",len(list_files))
	
	
	''' check exist directory & confirm about to delete it ''' 
	try:
		if(os.path.exists(path+keyword)):
			confirmToDelete = messagebox.askyesno("Directory Exist","Directory already exist with given name. Would you like to clean up before copy files in it?")
			
			if(confirmToDelete):
				processingTextField.config(state=NORMAL)
				processingTextField.insert(END,"Deleting the files under '"+path+keyword+"'\n", 'success')
				processingTextField.config(state=DISABLED)
				
				shutil.rmtree(path+keyword)
				os.mkdir(path+keyword)
		else:
			processingTextField.config(state=NORMAL)
			processingTextField.insert(END,"Creating new directory '"+path+keyword+"'\n", 'success')
			processingTextField.config(state=DISABLED)
			
			os.mkdir(path+keyword)
			
	except Exception as e:
		processingTextField.config(state=NORMAL)
		processingTextField.insert(END,"System Error.. '"+str(e)+"'\n", 'error')
		processingTextField.config(state=DISABLED)
	
	
	''' process the files now '''
	for file in list_files:
		try:
			if(file.endswith('.doc') or (file.endswith('.docx'))):
			
				processingTextField.config(state=NORMAL)
				processingTextField.insert(END,"\nProcessing.. '"+file+"'\n")
				processingTextField.config(state=DISABLED)
				
				list_doc_files.append(file)
				text=docxpy.process(path+file)
				text=text.lower()
				if(keyword in text):
					processingTextField.config(state=NORMAL)
					processingTextField.insert(END,"Keyword ")
					processingTextField.config(state=DISABLED)
					
					processingTextField.config(state=NORMAL)
					processingTextField.insert(END,"'"+keyword+"'", 'success')
					processingTextField.config(state=DISABLED)
					
					processingTextField.config(state=NORMAL)
					processingTextField.insert(END," found in document ")
					processingTextField.config(state=DISABLED)
					
					processingTextField.config(state=NORMAL)
					processingTextField.insert(END,"'"+file+"'", 'success')
					processingTextField.config(state=DISABLED)
					
					processingTextField.config(state=NORMAL)
					processingTextField.insert(END,"\n")
					processingTextField.config(state=DISABLED)
					
					list_doc_filter.append(file)
					shutil.copy(path+file, path+keyword+"/"+file)
					
					processingTextField.config(state=NORMAL)
					processingTextField.insert(END,"File '"+file+"' copied to '"+path+keyword+"'\n")
					processingTextField.config(state=DISABLED)
		except Exception as e:
			print('Error', str(e))
			list_files_problem.append(file)
			processingTextField.config(state=NORMAL)
			processingTextField.insert(END,"Error while processing file '"+file+"'\n'"+str(e)+"'\n", 'error')
			processingTextField.config(state=DISABLED)


	processingTextField.config(state=NORMAL)
	processingTextField.insert(END,"\n\n######### Process Completed ###########\n")
	processingTextField.insert(END,"Total Document Files '"+str(len(list_doc_files))+"'\n", 'success')
	processingTextField.insert(END,"Total Files with relavant keywords '"+str(len(list_doc_filter))+"'\n", 'success')
	processingTextField.insert(END,"Total Files contains errors '"+str(len(list_files_problem))+"'\n", 'error')
	processingTextField.config(state=DISABLED)

	resultTextField.config(state=NORMAL)
	resultTextField.delete(1.0,END)
	resultTextField.insert(END,"Total Document Files '"+str(len(list_doc_files))+"'\n", 'success')
	resultTextField.insert(END,"Total Files with relavant keywords '"+str(len(list_doc_filter))+"'\n", 'success')
	resultTextField.insert(END,"Total Files contains errors '"+str(len(list_files_problem))+"'\n", 'error')
	
	resultTextField.insert(END,"\n")	
	resultTextField.insert(END, "Browse Directory\n\n", hyperlink.add(openDirectory))
	resultTextField.insert(END, "\n######### Filtered Documents ###########\n\n")
	
	for doc in list_doc_filter:
		resultTextField.insert(END, doc + "\n", hyperlink.add(openDocument))
	
	resultTextField.config(state=DISABLED)
	
	print("Total Doc Files:",len(list_doc_files))   
	print("Total Filtered Files:",len(list_doc_filter))  
	print("Total Files with problems:",len(list_files_problem))
	print(list_files_problem)


''' open document in their default application '''
def openDocument(tag):
	index = int(tag[6:])
	filepath = inputDirectoryPath.get() + '/' + inputKeywords.get() + '/' + list_doc_filter[index-1]
	print('Hyperlink tag', tag, filepath)
	
	if platform.system() == 'Darwin':       # macOS
		subprocess.call(('open', filepath))
	elif platform.system() == 'Windows':    # Windows
		os.startfile(filepath)
	else:                                   # linux variants
		subprocess.call(('xdg-open', filepath))


''' open directory '''
def openDirectory(tag):
	if platform.system() == 'Darwin':       # macOS
		subprocess.call(('open', inputDirectoryPath.get()))
	elif platform.system() == 'Windows':    # Windows
		os.startfile(inputDirectoryPath.get())
	else:                                   # linux variants
		subprocess.call(('xdg-open', inputDirectoryPath.get()))



list_doc_files = []
list_doc_filter = []
list_files_problem = []	


''' parent window setup '''
root=Tk()
root.state('zoomed')
root.title('Document Sorter')
root.resizable(width=False,height=False)
root.configure(bg="black")


''' result frame '''
frm=Frame(root,width=root.winfo_screenwidth()/2,height=root.winfo_screenheight())
frm.configure(bg="black")
frm.place(x=root.winfo_screenwidth()/2,y=55)


''' Title '''
title=Label(root,text="Resume Finder",font=('',30,'bold'),bg='black', fg='white')
title.pack()


''' directory path '''
directoryPathLabel=Label(root,text="Directory Path:",font=('Arial',15,'bold'),bg='black', fg='white')
directoryPathLabel.place(x=100,y=100)
inputDirectoryPath=Entry(root,font=('',15,'bold'),bd=1,width=20)
inputDirectoryPath.place(x=280,y=100)
browse_btn=Button(root,text="Browse",font=('',11,'bold'),bd=1,command=lambda:browse(inputDirectoryPath, processingText, resultText))
browse_btn.place(x=520,y=100)


''' search keywords '''
searchKeywordLabel=Label(root,text="Search Keywords:",font=('Arial',15,'bold'),bg='black',fg='white')
searchKeywordLabel.place(x=100,y=150)
inputKeywords=Entry(root,font=('',15,'bold'),bd=1,width=20)
inputKeywords.place(x=280,y=150)
searchKeywordHintLabel=Label(root,text="you can provide multiple keywords with comma separated, e.g. python,django,ds...",font=('Arial',8,'normal'),bg='black',fg='white')
searchKeywordHintLabel.place(x=100,y=180)


''' search options '''
var=StringVar()
var.set('different')
searchOptionLabel=Label(root,text="Do you want to search all keywords in single file OR in every different file?",font=('Arial',10,'normal'),bg='black',fg='white')
searchOptionLabel.place(x=100,y=220)
searchOption1=Radiobutton(text='Single File',variable=var,value='single', bg='black', borderwidth=2, padx=3, pady=3, fg='white', selectcolor='black')
searchOption2=Radiobutton(text='Different File',variable=var,value='different', bg='black', borderwidth=2, padx=3, pady=3, fg='white', selectcolor='black')
searchOption1.place(x=100,y=240)
searchOption2.place(x=200,y=240)


'''Submit button '''
sub_btn=Button(root,text="Submit ",font=('',11,'bold'),bd=1,command=lambda:submit(inputDirectoryPath,inputKeywords, processingText, resultText))
sub_btn.place(x=100,y=300)


''' Reset Button '''
reset_btn=Button(root,text="Refresh",font=('',11,'bold'),bd=1,command=lambda:reset(inputDirectoryPath,inputKeywords, processingText, resultText))
reset_btn.place(x=180,y=300)


''' Processing area '''
processingText = ScrolledText(root, height=17, width=60, border=2)
processingText.insert('end', 'Initialized successfully..\n')
processingText.tag_config("error", background="red", foreground="white")
processingText.tag_config("success", background="green", foreground="white")
processingText.config(state=DISABLED)
processingText.place(x=100,y=root.winfo_screenheight()/2)


''' Result area '''
resultText = ScrolledText(frm, height=35, width=70, border=2)
resultText.insert('end', 'Result will be dispay here\n')
resultText.tag_config("error", background="red", foreground="white")
resultText.tag_config("success", background="green", foreground="white")
resultText.config(state=DISABLED)
resultText.place(x=25,y=45)

''' configure result area to support hyperlinks '''
hyperlink = HyperlinkManager(resultText)


''' start main program '''
root.mainloop()
