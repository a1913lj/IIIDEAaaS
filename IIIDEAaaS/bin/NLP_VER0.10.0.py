import re
import os
import sys

from janome.analyzer import Analyzer
from janome.charfilter import UnicodeNormalizeCharFilter, RegexReplaceCharFilter
from janome.tokenizer import Tokenizer as JanomeTokenizer  # sumyのTokenizerと名前が被るため
from janome.tokenfilter import POSKeepFilter, ExtractAttributeFilter

 
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer

from chardet.universaldetector import UniversalDetector

from pysummarization.nlpbase.auto_abstractor import AutoAbstractor
from pysummarization.tokenizabledoc.mecab_tokenizer import MeCabTokenizer
from pysummarization.abstractabledoc.top_n_rank_abstractor import TopNRankAbstractor

def detect_character_code(file):
    """
    pathnameから該当するファイルの文字コードを判別して、文字コードを返す
 
    :param pathname: 文字コードを判別したいファイル
    :return: 文字コード
    """
    detector = UniversalDetector()
    with open(file, 'rb') as f:
        detector.reset()
        for line in f.readlines():
            detector.feed(line)
            if detector.done:
                break
        detector.close()
    return detector.result['encoding']

def mecab_document_summarize(document):
    #os.environ['MECABRC'] = 'C:\Program Files\MeCab\etc\mecabrc'
    
    '''
    https://software-data-mining.com/python%E3%83%A9%E3%82%A4%E3%83%96%E3%83%A9%E3%83%AApysummarization%E3%82%92%E7%94%A8%E3%81%84%E3%81%9Failstm%E3%81%AB%E3%82%88%E3%82%8B%E6%96%87%E6%9B%B8%E8%A6%81%E7%B4%84/
    Parameters
    ----------
    document : TYPE
        DESCRIPTION.

    Returns
    -------
    result_dict : TYPE
        DESCRIPTION.

    '''
    auto_abstractor = AutoAbstractor()
    auto_abstractor.tokenizable_doc = MeCabTokenizer()
    auto_abstractor.delimiter_list = ["。", "\n"]
    abstractable_doc = TopNRankAbstractor()
    result_dict = auto_abstractor.summarize(document, abstractable_doc)
    
    print('\n要約:')
    rst = ''
    for sentence in result_dict["summarize_result"]:
        print(sentence.strip())
        rst += sentence
    
    return result_dict, rst

def janome_document_summarize(document):
    # 形態素解析(単語単位に分割する)
    analyzer = Analyzer(
        char_filters=[UnicodeNormalizeCharFilter(), RegexReplaceCharFilter(r'[(\)「」、。]', ' ')], 
        tokenizer=JanomeTokenizer(),
        token_filters=[POSKeepFilter(['名詞', '形容詞', '副詞', '動詞']), ExtractAttributeFilter('base_form')]
    )
    
    text = re.findall("[^。]+。?", document.replace('\n', ''))
    corpus = [' '.join(analyzer.analyze(sentence)) + u'。' for sentence in text]
    parser = PlaintextParser.from_string(''.join(corpus), Tokenizer('japanese'))
    summarizer = LexRankSummarizer()
    summarizer.stop_words = [' ', '。','\n']
    N = int(len(corpus)/10*3)
    if N <=0 : N=3
    summary = summarizer(document=parser.document, sentences_count=N)
    
    rst = ''
    print('\n要約:')
    for sentence in summary:
        print(text[corpus.index(sentence.__str__())])
        rst += text[corpus.index(sentence.__str__())]
    return summary, rst

def rf(file):
    with open(file, "r", encoding=detect_character_code(file)) as f:
        contents = f.readlines()
    contents = ''.join(contents)
    return contents
    
        
def wf(file, txt):
    with open(file, "w", encoding = "utf_8") as f:
        f.write(txt)
        
def fn_start_document_summarize(file, Category):    
    contents = rf(file)
    
    print('原文:', Category)
    print(contents)
    if Category.lower() == 'MeCab'.lower():
        summary = mecab_document_summarize(contents)
    else:
        summary = janome_document_summarize(contents)
    return summary

def main():
    fn1 = ''
    fn2 = ''
    Category = 'MeCab'
    #print(sys.argv)
    for i in range(len(sys.argv)):
        if i==1: fn1 = sys.argv[i]
        if i==2: fn2 = sys.argv[i]
        if i==3: Category = sys.argv[i]
    
    file =''
    if fn1 == '':
        file = input("filename input->")
    else:
        file = fn1
    summary, rst = fn_start_document_summarize(file, Category)
    
    print(u'文書要約結果:', summary)        
    if fn2 == '':
        fn2 = str(os.getpid()) + '.txt'
    wf(fn2, rst)

if __name__ == '__main__':
    main()
    

    
    
   