# -*- coding: utf-8 -*-

import fs
import logging
import win32com.client as win32
import os.path
from pathlib import Path

class WordDocSearcher(object):

    def search_word_docs(self, document_directory, search_term, search_recursively, update_text_callback):

        update_text_callback(f'{"Recursively " if search_recursively else ""}Searching Word Documents in {document_directory} for {search_term}.')

        document_directory = Path(document_directory)

        doc_paths = self.__gather_doc_paths(document_directory,search_recursively,update_text_callback)

        docs_with_search_term, docs_without_search_term, docs_with_errors = self.__search_docs_for_search_term(doc_paths,search_term, update_text_callback)

        docs_with_search_term = "\n".join(docs_with_search_term)

        update_text_callback(docs_with_search_term, do_replace=True)

    def __gather_doc_paths(self, document_directory, search_recursively, update_text_callback):
        '''Gathers the paths to Word Documents in the given document directory.
        
        Arguments:
            document_directory {str} -- The document directory to search.
            search_recursively {bool} -- Whether or not to search the directory recursively.
        
        Returns:
            list -- The Word Document paths in the given document_directory.
        '''

        logging.info(f'Gathering Word Document Paths from {document_directory}')

        update_text_callback(f'Gathering Word Document Paths from {document_directory}')

        doc_paths = []

        glob_search_string = (
            '**/*.doc?'
            if search_recursively
            else
            '*.doc?'
        )

        for glob_match in fs.open_fs(document_directory.__str__()).glob(glob_search_string): 

            doc_path = (document_directory / glob_match.path[1:]).__str__()

            if '~$' in doc_path: continue

            doc_paths.append(doc_path)
            
        logging.debug(f'Gathered {len(doc_paths)} document paths.')

        update_text_callback(f'Gathered {len(doc_paths)} document paths.')

        return doc_paths

    def __search_docs_for_search_term(self, doc_paths, search_term, update_text_callback):
        '''Searches the Word Documents identified by the given doc paths, for the given search term.
        
        Arguments:
            doc_paths {list} -- The doc paths of the Word Documents to search.
            search_term {str} -- The search term to look for.
        
        Returns:
            list -- The documents who contain the search term.
            list -- The documents who don't contain the search term.
        '''

        logging.info(f'Searching Word Docs for {search_term}.')

        update_text_callback(f'Searching Word Docs for {search_term}.')

        docs_with_search_term = []

        docs_without_search_term = []

        docs_with_errors = {}

        # Get a handle on the Word Application.
        try:

            msword = win32.gencache.EnsureDispatch('Word.Application')

        except:

            msword = None

        if msword is None: 

            logging.warn('For an unknown reason, Word is unavailable. Process halted.')

            update_text_callback('For an unknown reason, Word is unavailable. Process halted.')

        else:

            # For each of the doc paths,
            for doc_index, doc_path in enumerate(doc_paths):

                logging.debug(f'Searching doc {doc_index + 1} out of {len(doc_paths)}: {doc_path}')

                update_text_callback(f'Searching doc {doc_index + 1} out of {len(doc_paths)}: {doc_path}\n')

                logging.debug('Backing up raw doc binary data so unexpected changes during the search process can be overwritten.')

                doc_data = open(doc_path, 'rb').read()

                try:

                    # Open the doc invisibly in Word
                    word_doc = msword.Documents.Open(doc_path, Visible = False) 

                    if word_doc is None:

                        logging.warn(f'The word_doc instance for {doc_path} is None!')

                        docs_with_errors[doc_path] = f'The word_doc instance for {doc_path} is None!'

                        continue

                except Exception as e:

                    logging.warn(f'Error while trying to open {doc_path}: {str(e)}')

                    docs_with_errors[doc_path] = str(e)

                    continue

                search_term_found = False

                sections = word_doc.Sections

                # Look for the search term in the Word Doc.
                for section in word_doc.Sections:

                    for paragraph in section.Range.Paragraphs:

                        if search_term in paragraph.Range.Text: search_term_found = True

                    for header in section.Headers: 

                        if search_term in header.Range.Text: search_term_found = True

                    for footer in section.Footers: 

                        if search_term in footer.Range.Text: search_term_found = True

                    if search_term_found: break

                # Save the doc path in the appropriate list
                if search_term_found: docs_with_search_term.append(doc_path)

                else: docs_without_search_term.append(doc_path)

                # Close the Word Document. NOTE: ABSOLUTELY NECESSARY
                word_doc.Close()

                logging.debug('Overwritting the document with the previously collected data to undo any unexpected changes due to searching.')

                with open(doc_path, 'wb') as out_file: out_file.write(doc_data)

        if msword is not None: 
            
            try:

                msword.Quit()

            except Exception as e:

                logging.warn(f'Error while trying to close Word: {str(e)}')
                
            logging.debug(f'Out of {len(doc_paths)} documents, {len(docs_with_search_term)} contained the search term and {len(docs_without_search_term)} did not. {len(docs_with_errors)} had errors while trying to open.')

            update_text_callback(f'Out of {len(doc_paths)} documents, {len(docs_with_search_term)} contained the search term and {len(docs_without_search_term)} did not. {len(docs_with_errors)} had errors while trying to open.')

        return docs_with_search_term, docs_without_search_term, docs_with_errors