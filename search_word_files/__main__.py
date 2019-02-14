# -*- coding: utf-8 -*-

import logging
from word_doc_searcher import WordDocSearcher
from primary_ui import PrimaryUI

if __name__ == "__main__":
    '''
    Main Method
    '''

    logging.basicConfig(level=logging.DEBUG)

    primary_ui = PrimaryUI()

    word_doc_searcher = WordDocSearcher()

    primary_ui.start(word_doc_searcher.search_word_docs)