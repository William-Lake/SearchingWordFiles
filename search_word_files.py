try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import sys
import os
import pathlib
from docx import Document
from tkinter import Tk
from tkinter import IntVar
from tkinter import END
from tkinter import INSERT
from tkinter import Button
from tkinter import Label
from tkinter import Entry
from tkinter import Checkbutton
from tkinter import N,S,E,W
from tkinter.filedialog import askdirectory
from tkinter import scrolledtext
from tkinter import messagebox
import re

window = Tk()

window.title('Searching Word Docs')

window.geometry('630x250')

window.resizable(False,False)

use_regex = IntVar(window)

dir_to_search = ''

search_term = ''

def get_docx_text(path):

    # http://etienned.github.io/posts/extract-text-from-word-docx-simply/
    WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    PARA = WORD_NAMESPACE + 'p'

    TEXT = WORD_NAMESPACE + 't'

    document = zipfile.ZipFile(path)

    xml_content = document.read('word/document.xml')

    document.close()

    tree = XML(xml_content)

    paragraphs = []

    for paragraph in tree.getiterator(PARA):

        texts = [node.text
                for node in paragraph.getiterator(TEXT)
                if node.text]

        if texts:

            paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)

def perform_search():

    txt_search_results.delete(1.0,END)

    document_list = []

    for path, subdirs, files in os.walk(dir_to_search): 

        for name in files:

            if os.path.splitext(os.path.join(path, name).upper())[1] == ".DOCX" or os.path.splitext(os.path.join(path, name).upper())[1] == ".DOC":
                
                document_list.append(os.path.join(path, name))

    documents_containing_search_term = []

    for document_path in document_list:

        document_text = get_docx_text(document_path)

        if use_regex.get() == 1:

            regex = r'{0}'.format(search_term)

            matches = re.findall(regex,document_text,re.RegexFlag.MULTILINE)

            for match in matches:

                documents_containing_search_term.append('{} : {}'.format(document_path,match))

        elif search_term in document_text: documents_containing_search_term.append(document_path)

    results = []

    for document_containing_search_term in documents_containing_search_term:

        results.append(document_containing_search_term.replace(dir_to_search,''))

    txt_search_results.insert(INSERT,'\n'.join(results))

def btn_browse_clicked():

    global dir_to_search

    dir_to_search = askdirectory(initialdir=pathlib.Path(__file__).parent)

    lbl_browse.config(text=dir_to_search)

def btn_search_clicked():

    '''
    Check if a directory was selected
    check if a search term was provided
    '''
    err_msg = ''

    global search_term

    search_term = txt_search_term.get()

    if len(dir_to_search.strip()) == 0: err_msg += 'You need to select a directory first.\n'

    if len(search_term.strip()) == 0: err_msg += 'You need to provide a search term.'

    if len(err_msg) == 0: perform_search()
    
    else: messagebox.showerror('Need Info',err_msg)

# --- BUILDING UI ---

btn_browse = Button(window,text='Browse',command=btn_browse_clicked)

btn_browse.grid(column=0,row=0)

lbl_browse = Label(window)

lbl_browse.grid(column=1,row=0)

lbl_search_term = Label(window,text='Search Term:')

lbl_search_term.grid(column=0,row=1,stick=W)

txt_search_term = Entry(window,width=80)

txt_search_term.grid(column=1,row=1,sticky=N+S+W)

btn_regex = Checkbutton(window,variable=use_regex,text="Regex?")

btn_regex.grid(column=2,row=1,stick=E)

txt_search_results = scrolledtext.ScrolledText(window,width=75,height=10)

txt_search_results.grid(column=0,row=2,columnspan=3)

btn_search = Button(window,text='Search',command=btn_search_clicked)

btn_search.grid(column=2,row=3)

window.mainloop()