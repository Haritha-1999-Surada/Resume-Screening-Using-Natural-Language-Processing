from tkinter import *
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import re
import os
import io
import glob
from pathlib import Path
import spacy
import docx2txt
import PyPDF2
import csv
import pandas as pd
import nltk
nltk.download('stopwords')
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('wordnet')
from nltk.corpus import wordnet
from itertools import islice
# Load English tokenizer, tagger, parser and NER
nlp = spacy.load('en_core_web_sm')
file_name_dict={}
file_mail_dict={}
file_ph_dict={}
file_exp_dict={}
file_skill_dict={}
result_dict = {}

#function which the file extensions and if it is docx or pdf it calls the functions to convert them into txt format
def convertMultiple(docDir, txtDir):
    file_list=[]
    for doc in os.listdir(docDir): 
        doc_len=len(doc.split("."))
        fileExtension = doc.split(".")[-1]
        dname=doc.split(".")[0]
        print(dname)
        if fileExtension == "docx":
            #function to convert a docx file to a txt file
            docFilename = docDir + str(doc) 
            text = docx2txt.process(docFilename) 
            textFilename = txtDir + dname + ".txt"
            textFile = open(textFilename, "w", encoding="UTF-8") 
            textFile.write(text)
            textFile.close()
            
        elif fileExtension == "pdf":
            docFilename = docDir + str(doc) 
            textFilename = txtDir + dname + ".txt"
            text = pdftotext(docFilename,textFilename) 
            
            
        elif fileExtension =="txt":
            docFilename = docDir + str(doc) 
            textFilename = txtDir + dname + ".txt"
            textFile = open(textFilename, "w", encoding="UTF-8") 
            textFile.write(text)
            textFile.close()

#function to convert a pdf file to a txt file 
def pdftotext(file,textFilename):
    content = ""
    p1 = open(file, "rb")
    pdf = PyPDF2.PdfFileReader(p1)
    num=pdf.getNumPages()
    for i in range(0, num):
        content += pdf.getPage(i).extractText() + "\n"
    content = " ".join(content.replace(u"\xa0", u" ").strip().split())
    new=""
    for i in content:
        new+=i
    textFile = open(textFilename, "w", encoding="UTF-8") 
    textFile.write(new)
    textFile.close()
    p1.close()

#function to find e-mail ids
def email(newDir):
    path = newDir + "*.txt"
    path = path.replace(os.sep, '/')
    text_files = glob.glob(path)
    files = set(text_files)
    files = list(files)
    for each in files:
        filename=each.split('\\')[-1]
        filename=filename.split('.')[0]
        #print(filename)                    
        with open(each, 'r', encoding='utf8', errors='ignore') as f:
            search = f.read()
        email = None           
        pattern = re.compile('[\w\.-]+@[\w\.-]+')
        matches = pattern.findall(search)
        email = matches
        email = set(email)
        email = list(email)
        new = "" 
  
      
        for x in email: 
            new += x+" , " 
        print(new)    
        file_mail_dict[filename]=new

#function to find phone numbers
def phone(file):
  path = newDir + "*.txt"
  path = path.replace(os.sep, '/')
  text_files = glob.glob(path)
  files = set(text_files)
  files = list(files)
  for each in files:
    filename=each.split('\\')[-1]
    filename=filename.split('.')[0]
    phone = None
    with open(each, 'r', encoding='utf8', errors='ignore') as f:
      search = f.read()
      phone = re.search(r'[6789][0-9]{9}',search)
    print(phone[0])
    file_ph_dict[filename]=phone[0]

import spacy
from spacy.matcher import Matcher
import pandas as pd

# load pre-trained model
nlp = spacy.load('en_core_web_sm')

# initialize matcher with a vocab
matcher = Matcher(nlp.vocab)

#function to find the skillsets
def skillsets(file1):
  path = newDir + "*.txt"
  path = path.replace(os.sep, '/')
  text_files = glob.glob(path)
  files = set(text_files)
  files = list(files)
  data = pd.read_csv(file1)
  for each in files:
    filename=each.split('\\')[-1]
    filename=filename.split('.')[0]
    extension=each.split('.')[-1]
    if extension == 'txt':
      skillsets=set()
      
      with open(each, 'r', encoding='utf8', errors='ignore') as f:
        search = f.read().lower().split()
        search=' '.join(search)
        # removing stop words and implementing word tokenization
        nlp_text = nlp(search)
        tokens = [token.text for token in nlp_text if not token.is_stop]
        # extract values
        skills=[]
        for i in list(data.columns.values):
          skills.append(i.lower())
        for token in tokens:
          if ',' in token:
            tokens.extend(token.split(','))
          if '(' in token:
            tokens.extend(token.split('('))
        for token in tokens:
          if token.lower() in skills:
            skillsets.add(token)
        
       
        # check for bi-grams and tri-grams
        for token in nlp_text.noun_chunks:
          token = token.text.lower().strip()
          if token in skills:
            skillsets.add(token)
        
    new=""        
    skillsets=list(skillsets)
    for x in skillsets:
        new += x+" , " 
    new=tokenize1(new)    
    print(new)
    file_skill_dict[filename]=new

#function to remove phone numbers for finding the name and experience
def remove_phone(file):
    with open(file, 'r', encoding='utf8', errors='ignore') as f:
        search = f.read()
    try:
    
        phone = None           
        pattern = re.compile(r'([+(]?\d+[)\-]?[ \t\r\f\v]*[(]?\d{2,}[()\-]?[ \t\r\f\v]*\d{2,}[()\-]?[ \t\r\f\v]*\d*[ \t\r\f\v]*\d*[ \t\r\f\v]*)')
        matches = pattern.findall(search)
        matches = [re.sub(r'[,.]', '', el) for el in matches if len(re.sub(r'[()\-.,\s+]', '', el))>6]
        matches = [re.sub(r'\D$', '', el).strip() for el in matches]
        matches = [el for el in matches if len(re.sub(r'\D','',el)) <= 15]
        for el in list(matches):
            if len(el.split('-')) > 3: continue
            for x in el.split("-"):
                try:
                    if x.strip()[-4:].isdigit():
                        if int(x.strip()[-4:]) in range(1900, 2100):
                            matches.remove(el)
                except: pass                  
        phone = matches
        phone = set(phone)
        phone = list(phone)
        search = ""
        with open(file, 'r', encoding='utf8', errors='ignore') as f:
                search = f.read()
                for i in phone:
                    search = search.replace(i," ")
        with open(file, "w",encoding='utf8', errors='ignore') as f:
            f.write(search)
    except:
        pass

#function to remove e-mails  for finding the name and experience
def remove_email(file):
    with open(file, 'r', encoding='utf8', errors='ignore') as f:
        search = f.read()
    try:    
        email = None           
        pattern = re.compile('[\w\.-]+@[\w\.-]+')
        matches = pattern.findall(search)
        email = matches
        email = set(email)
        email = list(email)
        search = ""
        with open(file, 'r', encoding='utf8', errors='ignore') as f:
                search = f.read()
                for i in email:
                    search = search.replace(i," ")
        with open(file, "w",encoding='utf8', errors='ignore') as f:
            f.write(search)
    except:
        pass

#function to find the name
def name(file):
    filename=file.split('\\')[-1]
    filename=filename.split('.')[0]
    with open(file, 'r', encoding='utf8', errors='ignore') as f:
        search = f.read().upper()
           
    Sentences = nltk.sent_tokenize(search)
    Tokens = []
    for Sent in Sentences:
        Tokens.append(nltk.word_tokenize(Sent)) 
    Words_List = [nltk.pos_tag(Token) for Token in Tokens]


    Nouns_List = []
    
    for List in Words_List:
        for Word in List:
            if re.match('[NN.*]', Word[1]):
                Nouns_List.append(Word[0])
  
    text=' '.join(map(str,Nouns_List))
  
    nlp_text = nlp(text)
    
    
    # First name and Last name are always Proper Nouns
    pattern = [{'POS': 'PROPN'}, {'POS': 'PROPN'}]
    
    matcher.add('NAME', [pattern])
    
    matches = matcher(nlp_text)
    for match_id, start, end in matches:
        span = nlp_text[start:end]
        #print("Name:",span.text)
        break
    
    
    
    file_name_dict[filename]=span.text

#function to find the experience
def exp(file):

    filename=file.split('\\')[-1]
    filename=filename.split('.')[0]
    with open(file, 'r', encoding='utf8', errors='ignore') as f:
        search = f.read()
    try: 
    #print(search)
        new_left1 = []
        new_left2 = []
        left = 0
        
        left2 = re.search(r"(?:[a-zA-Z'-]+[^a-zA-Z'-]+){0,7}experience(?:[^a-zA-Z'-]+[a-zA-Z'-]+){0,7}", search)
        if(left2 != None):
            left = left2.group()
        left4 = re.search(r"(?:[a-zA-Z'-]+[^a-zA-Z'-]+){0,2}years(?:[^a-zA-Z'-]+[a-zA-Z'-]+){0,2}", search)
        if(left4 != None):
            left = left4.group()
        left5 = re.search(r"(?:[a-zA-Z'-]+[^a-zA-Z'-]+){0,2}year(?:[^a-zA-Z'-]+[a-zA-Z'-]+){0,2}", search)
        if(left5 != None):
            left = left5.group()    
    
        if(left == 0):
            print("Experience : ", 0)
            return
    #print(left)
        left1 = re.findall('[0-9]{1,2}',left)
        left1_int = list(map(int, left1))
        #print(left1_int)
        for a in left1:
            for b in left1_int:
                if len(a) <= 2 and b < 30:
                    new_left1.append(a)
        left1 = ''.join(new_left1[0])
        left2 = re.findall('[0-9]{1,2}.[0-9]{1,2}',left)
    
        exp = []
        if not left2:
            exp.append(left1)
            exp = ''.join(left1)
        else:
            exp.append(left2)
            exp = ''.join(left2)
    except:
        exp=0
    print("Experience : ", exp)
    file_exp_dict[filename]=exp

#function to call the functions which finds experience and name
def name_and_exp_call():
    path = newDir + "*.txt"
    path = path.replace(os.sep, '/')
    text_files = glob.glob(path)
    files = set(text_files)
    files = list(files)
    #print(files)
    for file in files:
        remove_phone(file)
        remove_email(file)
        name(file)
        exp(file)
        print("\n")

#function to store the extracted info in a csv
def store_csv():
    with open(r"C:\Users\JA\Resume Screening\Details.csv",'w',encoding='utf8', errors='ignore') as fd:
        writer=csv.writer(fd)
        writer.writerow(["ID","Name","Experience","Skillset","Phone no","EmailID"])
        for key in file_name_dict.keys():
            writer.writerow([key,file_name_dict[key],file_exp_dict.get(key,0),file_skill_dict[key],file_ph_dict[key],file_mail_dict[key]])

from difflib import SequenceMatcher

#function to calculate score and store the final result in a csv
def score(file2):
  import csv
  import math
  
  count_scr=0
  score=0
  maximum=0

  result_dict = {}
  #open the resume csv which has the job seeker's info like name,experience,etc
  with open(r"C:\Users\JA\Resume Screening\Details.csv",'r') as resume_csv_data:
    csv_data = csv.DictReader(resume_csv_data)
    for resume_row in csv_data:
      ID = resume_row["ID"]
      experience = resume_row["Experience"]
      skills= tokenize1(resume_row["Skillset"]).split(",")
      #open the resume csv which has the companies info like job description,experience required,etc
      job_csv = csv.DictReader(open(file2, mode='r'))
      for i,job_row in enumerate(job_csv):
        job_experience_required=job_row["experience"]
        job_description=job_row["decrption"]
        #job_ID = job_row["company"]
        job_description=tokenize1(job_description).split(',')
        for skill in skills:
          for des in job_description:
            similarity=SequenceMatcher(None,skill,des).ratio()
            if similarity > maximum:
              maximum=similarity
            if similarity >=0.9999999999:
              count_scr=count_scr+1
          score=score+maximum
#score is calculated by adding the maximum similarity skill with the experience of the job seeker and if any skill is exactly 
#matched it is multiplied by 0.2 and added to the final score
        score=score+float(experience)+(count_scr*0.2)    
        print("ID:{} Score: {}".format(ID,score))
        if result_dict.get(ID,0)==0:
          result_dict[ID]=(score)
        else:
          result_dict[ID].append(score)
        score=0
        count_scr=0
        maximum=0
  
   
  
  with open(r"C:\Users\JA\Resume Screening\Score.csv",'w',encoding='utf8', errors='ignore') as fd:
    writer=csv.writer(fd)
    writer.writerow(["EmailID","EmployeeID/Score"])
    for key in file_mail_dict.keys():
      #print([file_mail_dict[key],result_dict[key]])
      writer.writerow([file_mail_dict[key],result_dict[key]])

#function to pre-process the data
def tokenize(string, newFilename):

    doc = nlp(string)
    str1 = '\n'
    for word in (doc.sents): #used for sentence tokenization
        sentence = nlp(word.text) #used for text tokenization

        for token in sentence:
            #checks whether the word is a stop word , pronoun or a punctuation
            if token.is_stop!=True and token.is_punct != True and token.pos_ != 'PRON':  
                
                str1 = str1 + str(token.lemma_) #stores the lemmatized word in the variable str1
                with open(newFilename, "a+", encoding = "UTF-8") as f:

                    f.write(str1)
                    str1 = ' '
        str1 = '\n'

#function to refine the skillsets
def tokenize1(string):
    doc = nlp(string)
    str1 = ' '
    for word in (doc.sents):
        sentence = nlp(word.text)

        for token in sentence:
            if token.is_stop!=True and token.is_punct != True and token.pos_ != 'PRON':
                #print(token)
                
                
                str1 = str1 + str(token.lemma_)+","
              
    return str1

#function to check whether the file is in txt format
def parse(txtDir):
    path = txtDir + "*.txt"
    path = path.replace(os.sep, '/')
    text_files = glob.glob(path)
    files = set(text_files)
    files = list(files)   
    for each in files:
        each = each.replace(os.sep, '/')
        each = os.path.basename(each)
        
        
        tname, extension = each.split(".")
        newFilename = newDir + str(tname) + ".txt"
        if extension == "txt":
            textFilename = txtDir + str(each) 
        with open(textFilename, 'r', encoding="UTF-8") as f:
            string = f.read().lower()
            tokenize(string, newFilename)

#function to remove the files - the converted files and the preprocessed files
def remove_text():
  d='C:\\Users\\JA\\Resume Screening\\Text\\' 
  filesToRemove = [os.path.join(d,f) for f in os.listdir(d)]
  for f in filesToRemove:
    os.remove(f)
    d='C:\\Users\\JA\\Resume Screening\\New\\'
    filesToRemove = [os.path.join(d,f) for f in os.listdir(d)]
    for f in filesToRemove:
      os.remove(f)

#interface
root=Tk()
root.geometry("1000x500")
root.title("RESUME SCREENING")
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

class FolderSelect(Frame):
    def __init__(self,parent=None,folderDescription="",**kw):
        Frame.__init__(self,master=parent,**kw)
        self.folderPath = StringVar()
        self.lblName = Label(self, text=folderDescription)
        self.lblName.grid(row=0,column=0)
        self.entPath = Entry(self, textvariable=self.folderPath)
        self.entPath.grid(row=0,column=1)
        self.btnFind = ttk.Button(self, text="Browse Folder",command=self.setFolderPath)
        self.btnFind.grid(row=0,column=2)
    def setFolderPath(self):
        folder_selected = filedialog.askdirectory()
        self.folderPath.set(folder_selected)
    @property
    def folder_path(self):
        return self.folderPath.get()

class FileSelect(Frame):
    def __init__(self,parent=None,fileDescription="",**kw):
        Frame.__init__(self,master=parent,**kw)
        self.filePath = StringVar()
        self.lblName = Label(self, text=fileDescription)
        self.lblName.grid(row=0,column=0)
        self.entPath = Entry(self, textvariable=self.filePath)
        self.entPath.grid(row=0,column=1)
        self.btnFind = ttk.Button(self, text="Browse File",command=self.setFilePath)
        self.btnFind.grid(row=0,column=2)
    def setFilePath(self):
        file_selected = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select a File",
                                          filetypes = (("Text files",
                                                        "*.txt*"),("Pdf files","*.pdf*"),("Document files","*.docx*"),
                                                       ("all files",
                                                        "*.*")))
        self.filePath.set(file_selected)
    @property
    def file_path(self):
        return self.filePath.get()

def doStuff():
    global folder1
    global docDir
    global newDir
    global txtDir
    global label_file
    global output_file
    global file1
    global file2
    folder1 = directory1Select.folder_path
    file1=file1Select.file_path
    file2=file2Select.file_path
    docDir = folder1+'/'
    l = folder1.split('/')
    l.pop()
    pat = '/'.join(l)
    os.chdir(pat)
    os.system('mkdir Text')
    os.system('mkdir New')
    txtDir = pat+'/Text/'
    newDir = pat+'/New/'
    print("Doing stuff with folder", folder1)
    convertMultiple(docDir, txtDir)
    parse(txtDir)
    email(newDir)
    phone(newDir)
    skillsets(file1)
    name_and_exp_call()
    store_csv()
    score(file2)
    remove_text()
    output_file=pat+'/Score.csv'
    label_file["text"]='Your Output File Was  Already Saved In Following Path \n'+output_file+ '\n You can view the output by clicking on the button below '


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path =output_file
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name

    df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None

def clear_data():
    tv1.delete(*tv1.get_children())
    return None
wrapper1=   LabelFrame(root,text="UPLOAD RESUMES ")
wrapper2=   LabelFrame(root,text="OUTPUT PATH")
wrapper3=   LabelFrame(root,text="OUTPUT DATA")

wrapper1.pack(fill="both" ,expand="yes" ,padx=20 ,pady=10)
wrapper2.pack(fill="both" ,expand="yes" ,padx=20 ,pady=10)
wrapper3.pack(fill="both" ,expand="yes" ,padx=20 ,pady=10)

folderPath = StringVar()
skillsPath=StringVar()
jdpath=StringVar()

folder1 = str()
newDir= str()
txtDir=str()
docDir=str()
output_file=str()
file1=str()
file2=str()

directory1Select = FolderSelect(wrapper1,"Upload Resumes Folder")
directory1Select.grid(row=0)
file1Select = FileSelect(wrapper1,"Select Skills File")
file1Select.grid(row=1)
file2Select = FileSelect(wrapper1,"Select Job Description File")
file2Select.grid(row=2)

c = ttk.Button(wrapper1, text="Screen Resumes", command=doStuff)
c.grid(row=3,column=0)


# Button
button2 = tk.Button(wrapper2, text="SHOW OUTPUT", command=lambda: Load_excel_data())
button2.place(rely=0.65, relx=0.30)

# The file/file path text
label_file = ttk.Label(wrapper2, text=" RELAX!! Resumes Are Getting Screened,PLEASE WAIT... :)")
label_file.place(rely=0, relx=0)


## Treeview Widget
tv1 = ttk.Treeview(wrapper3)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(wrapper3, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(wrapper3, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget


root.mainloop()




