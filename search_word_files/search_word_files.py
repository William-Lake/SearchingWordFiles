import logging
from word_doc_searcher import WordDocSearcher
from primary_ui import PrimaryUI

def write_out_results(docs_with_search_term, docs_without_search_term, search_term):
    '''Writes the results of the Word Document search to the file Results.txt in the local directory.
    
    Arguments:
        docs_with_search_term {list} -- The paths of the docs who contained the search term.
        docs_without_search_term {list} -- The paths of the docs who did not contain the search term.
        search_term {str} -- The search term originally provided.
    '''

    logging.info('Writing out results.')

    with open('Docs_With_Search_Term.txt','w+') as out_file: out_file.write('\n'.join(docs_with_search_term))

    with open('Docs_Without_Search_Term.txt','w+') as out_file: out_file.write('\n'.join(docs_without_search_term))

def main():
    '''
    Main Method
    '''

    logging.basicConfig(level=logging.DEBUG)

    primary_ui = PrimaryUI()

    word_doc_searcher = WordDocSearcher()

    primary_ui.start(word_doc_searcher.search_word_docs)