from word_doc_searcher import WordDocSearcher
from primary_ui import PrimaryUI

def main():

    primary_ui = PrimaryUI()

    word_doc_searcher = WordDocSearcher()

    primary_ui.start(word_doc_searcher.search_word_docs)
