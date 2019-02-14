# -*- coding: utf-8 -*-

import PySimpleGUI as sg

class PrimaryUI(sg.Window):

    DEFAULT_UPDATE_TEXT = 'Results....'

    SEARCH = 'Search'

    SAVE = 'Save'

    RESET = 'Reset'

    def __init__(self):

        super().__init__('Search Word Docs')

        self.create_layout()
        
    def create_layout(self):

        self.txt_document_directory = sg.InputText(do_not_clear = True)

        btn_browse = sg.FolderBrowse()

        self.txt_search_term = sg.InputText(do_not_clear=True)

        self.chk_recursive = sg.Checkbox('Recursive Searching',default=True)

        btn_search = sg.Button(self.SEARCH)

        self.txt_updates = sg.Multiline(self.DEFAULT_UPDATE_TEXT, size=(50,15))

        self.btn_save = sg.Button(self.SAVE, disabled=True)

        self.btn_clear = sg.Button(self.RESET)

        layout = [
            [btn_browse, self.txt_document_directory],
            [self.txt_search_term],
            [btn_search, self.chk_recursive],
            [self.txt_updates],
            [self.btn_save,self.btn_clear]
        ]

        self.Layout(layout)

    def start(self, callback):

        while True:

            event, values = self.Read()

            if event is None: break

            elif event == self.SEARCH:

                if self.data_valid(): 
                    
                    self.txt_updates.Update(self.DEFAULT_UPDATE_TEXT)
                    
                    self.execute_callback(callback)

                    self.btn_save.Update(disabled=False)

            elif event == self.SAVE:

                file_paths = values[3].strip()

                if file_paths:

                    save_file_path = sg.PopupGetFile('Save Results As...', save_as=True, file_types=(('Text Files', '*.txt'),), no_window = True)

                    with open(save_file_path,'w+') as out_file: out_file.write(file_paths)

                self.btn_save.Update(disabled=True)

            elif event == self.RESET:

                self.txt_updates.Update(self.DEFAULT_UPDATE_TEXT)

                self.btn_save.Update(disabled=True)

        self.Close()

    def data_valid(self):

        self.document_directory = self.txt_document_directory.Get()

        self.search_term = self.txt_search_term.Get()

        self.search_recursively = self.chk_recursive.Get() == 1

        return (
            (self.document_directory is not None and len(self.document_directory.strip()) > 0) and
            (self.search_term is not None and len(self.search_term.strip()) > 0)
        )

    def execute_callback(self,callback):

        callback(
            self.document_directory, 
            self.search_term, 
            self.search_recursively,
            self.update_status_text
            )

        self.document_directory = None

        self.search_term = None

        self.search_recursively = None

    def update_status_text(self,text,do_replace=False):

        if do_replace: update_text = text

        else:

            update_text = (
                text
                if self.txt_updates.Get() == self.DEFAULT_UPDATE_TEXT
                else
                f'{text}\n{self.txt_updates.Get()}'
            )

        self.txt_updates.Update(update_text)

        self.Refresh()
