import pandas as pd
import numpy as np
import requests
import re
import os
from bs4 import BeautifulSoup
import PyPDF2
import urllib
import docx2txt
import win32com.client
from pdf2image import convert_from_path
from pytesseract import image_to_string
from PIL import Image
import pytesseract
import nltk
import string
import log_mod

#########################################################

lg = log_mod.log_help()

#########################################################

def url_retrieve_list(url):
    try:
        req = requests.get(url)
        html_doc = req.text
        soup = BeautifulSoup(html_doc , "lxml")
        a_tags = soup.find_all('a')
        base = r"https://github.com"
        list= []
        for link in a_tags:
            temp = link.get('href')
            if temp.endswith(".pdf")or temp.endswith(".docx") or temp.endswith(".doc") or temp.endswith(".rtf"):
                ext = temp.rsplit(".", 1)[1]
                list.append(base+temp.replace("blob","raw"))

        return list
    except Exception as e:
        lg.log("error while using url_retrieve_list function"+"\n"+e, _type = 'error')
        
##################################################################   

def download_files(URL):
    base = os.getcwd()
    if not os.path.isdir(os.getcwd()+"\\Files"):
        os.mkdir(os.getcwd()+"\\Files")
    else:
        pass
    if get_name(URL).endswith(".doc") or get_name(URL).endswith(".rtf"):
        word = win32com.client.Dispatch("Word.Application")
        word.visible = False
        wb = word.Documents.Open(URL)
        New_name = get_name(URL).rsplit(".")[0]+".docx"
        doc_path = os.getcwd()+"\\Files\\"+New_name
        wb.SaveAs(doc_path,FileFormat=16)
        
    else:
       
        doc_path = os.getcwd()+"\\Files\\"+helper.get_name(URL)
        response = urllib.request.urlretrieve(URL, doc_path)
    
    return doc_path

##################################################################

def get_text_OCR(pdf_file, poppler_path, tesseract_path):
    try:
        images = convert_from_path(pdf_file,poppler_path=poppler_path)
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        final_text = ""
        for pg, img in enumerate(images):
            final_text += image_to_string(img)
        return final_text
    except Exception as e:
        lg.log("error while using get_text_OCR function"+"\n"+e, _type = 'error')
  
   #######################################################################     

def read_file_pdf(path, poppler_path, tesseract_path):
    
    try:
        '''
        Text from the True PDFs
        
        '''
        f = PyPDF2.PdfFileReader(path)
        no_page = f.getNumPages()
        text = ""
        for i in range(no_page):
            text = text + f.getPage(i).extractText()
        length_txt = len(text)
        '''
        Text from Images
        '''
        img_text = get_text_OCR(path, poppler_path, tesseract_path)
        
        '''
        Whichever has more text tength will be returned
        
        '''
        if(length_txt>len(img_text)):
            return text
        else:
            return img_text
        
    except Exception as e:
        lg.log("error while using read_file_pdf function"+"\n"+e, _type = 'error')
        
############################################################################

def read_file_docx(path):
    try:
        text = docx2txt.process(path)
        return text
    except Exception as e:
        lg.log("error while using read_file_docx function"+"\n"+e, _type = 'error')
        
#################################################################################


def get_data_dict(url, poppler_path, tesseract_path):
    
    path = download_files(url)
#     response = urllib.request.urlretrieve(url, path)
    
    ext = url.rsplit(".", 1)[1]
    if ext=="pdf":
        try:
            txt = read_file_pdf(path, poppler_path, tesseract_path)
            file_name = get_name(url)
            emails = find_email(txt)
            git_link = find_git(txt)
            linkedin = find_linkedin(txt)
            found_skills = tuple(find_skills(txt))
            return (file_name,emails,git_link,linkedin,found_skills)
        except Exception as e:
            lg.log("error while reading" +str(file_name)+"\n"+e, _type = 'error')
        
    else:
        try:
            txt = read_file_docx(path)
            file_name = get_name(url)
            emails = find_email(txt)
            git_link = find_git(txt)
            linkedin = find_linkedin(txt)
            found_skills = tuple(find_skills(txt))
            os.system("taskkill /f /im  WINWORD.EXE")
            return (file_name,emails,git_link,linkedin, found_skills)
    
        except Exception as e:
            lg.log("error while reading" +str(file_name)+"\n"+e, _type = 'error')

###########################################################################

def find_email(txt):
    try:
        email = re.findall("[^\s]{1}[a-zA-Z0-9\._]+[@]{1}[a-z]+[\.]{1}com", txt)
        email = list(set(email))
        
        if len(email)>0:
            return email
        else:
            return None

    except Exception as e:
        lg.log("error while using find_email function"+"\n"+e, _type = 'error')

############################################################################

def find_git(txt):
    try:
        git_link = re.findall("https:\/\/github\.com\/[a-zA-Z0-9]+", txt)
        git_link = git_link+re.findall("HTTPS:\/\/GITHUB\.COM\/[a-zA-Z0-9]+", txt)
        git_link = list(set(git_link))
        
        if len(git_link)>0:
            return git_link
        else:
            return None
        
    except Exception as e:
        lg.log("error while using find_git function"+"\n"+e, _type = 'error')
        
############################################################################
def find_linkedin(txt):
    try:
        linkedin = re.findall("https:\/\/www\.linkedin\.com\/[a-zA-Z]{2,3}\/.+\/", txt)
        linkedin = linkedin+re.findall("HTTPS:\/\/GITHUB\.COM\/[a-zA-Z]{2,3}\/.+\/", txt)
        linkedin = linkedin+re.findall("www\.linkedin\.com\/[a-zA-Z]{2,3}\/.+\S", txt)
        linkedin = list(set(linkedin))
        
        if len(linkedin)>0:
            return linkedin
        else:
            return None
        
    except Exception as e:
        lg.log("error while using find_linkedin function"+"\n"+e, _type = 'error')

############################################################################
def get_name(url):
    try:
        file = url.rsplit("/", maxsplit=1)[1]
        file = file.replace("%2B", "+")
        file = file.replace("%20", " ")
        return file
    except Exception as e:
        lg.log("error while using get_name function"+"\n"+e, _type = 'error')


################################  Creating a Skills Database #################################################

df = pd.read_excel("Tech_Skills.xlsx")

#changing each word to lowercase
df["skills"] = df.skills.apply(lambda x: x.lower())
        
# creating a set of Skills
skills = set(df["skills"])

#############################################################################

def find_skills(txt):
    try:
        
        # Replacing the Tabs and New characters
        txt = txt.replace('\t', ' ')
        txt = txt.replace('\n', ' ')

        # Tokenizing
        stop_words = set(nltk.corpus.stopwords.words('english'))
        word_tokens = nltk.tokenize.word_tokenize(txt)

        #removing punct
        word_tokens = [x.lower() for x in word_tokens if x not in string.punctuation]

        #removing stopwords
        word_tokens = [x for x in word_tokens if x not in stop_words]
        
        # removing string which don't have alphabets
        word_tokens = [x for x in word_tokens if re.search('[a-zA-Z]+',x)]
        
        #bigrams
        bigrm = nltk.bigrams(word_tokens)
        bigrm = [*map(' '.join, bigrm)]

        #trigrams
        trigrm = nltk.trigrams(word_tokens)
        trigrm = [*map(' '.join, trigrm)]
        
        #making a final list
        data = bigrm+word_tokens+trigrm

        # List of found Skills
        found_skills = set([x for x in data if x in skills])

        return found_skills
    except Exception as e:
        lg.log("error while using find_skills function"+"\n"+e, _type = 'error')