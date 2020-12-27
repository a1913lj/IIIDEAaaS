# -*- coding: utf-8 -*-
"""
Created on Fri Nov 20 22:13:30 2020

@author: Jiang.Lijin
"""
#https://software-data-mining.com/python%E3%81%AB%E3%82%88%E3%82%8B%E8%87%AA%E7%84%B6%E8%A8%80%E8%AA%9E%E5%87%A6%E7%90%86%E6%8A%80%E6%B3%95%E3%82%92%E3%81%B5%E3%82%93%E3%81%A0%E3%82%93%E3%81%AB%E4%BD%BF%E7%94%A8%E3%81%97%E3%81%9F/
#https://github.com/miso-belica/sumy
#https://deepblue-ts.co.jp/python/sumy-extractives-ummarization/


#pip install sumy tinysegmenter Janome
#pip install -U spacy
#pip install neologdn
#pip install emoji --upgrade
#pip install mojimoji
#pip install -U ginza

import os
import sys
import re
import codecs
from janome.analyzer import Analyzer
from janome.charfilter import UnicodeNormalizeCharFilter, RegexReplaceCharFilter
from janome.tokenizer import Tokenizer as JanomeTokenizer  # sumyのTokenizerと名前が被るため
from janome.tokenfilter import POSKeepFilter, ExtractAttributeFilter
from chardet.universaldetector import UniversalDetector

from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer
    
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

def clean(text):
    text = text.replace("\n", "")
    text = text.replace("\u3000", "")
    text = text.replace("「", "")
    text = text.replace("」", "")
    text = text.replace("（", "")
    text = text.replace("）", "")
    text = text.replace("、", "")
    text = text.replace("。", "")
    text = text.replace(" ", "")
    return text

def rd(file=''):
    if not len(file):
        file = input("filename input->")
    
    with codecs.open(file, 'r', detect_character_code(file)) as f:
        contents = f.readlines()
    contents = ''.join(contents)
    
    text = re.findall("[^。]+。?", contents.replace('\n', ''))
    
    return text

def wf(file, txt):
    with open(file, "w", encoding = "utf_8") as f:
        f.write(txt)
    
def fn_start_document_summarize(text):  
    # 形態素解析(単語単位に分割する)
    tokenizer = JanomeTokenizer('japanese')
    char_filters=[UnicodeNormalizeCharFilter(), RegexReplaceCharFilter(r'[(\)「」、。]', ' ')]
    token_filters=[POSKeepFilter(['名詞', '形容詞', '副詞', '動詞']), ExtractAttributeFilter('base_form')]
    
    analyzer = Analyzer(
        char_filters=char_filters,
        tokenizer=tokenizer,
        token_filters=token_filters
    )
 
    corpus = [' '.join(analyzer.analyze(sentence)) + u'。' for sentence in text]
    #print(corpus, len(corpus))
    
    # 文書要約処理実行    
    parser = PlaintextParser.from_string(''.join(corpus), Tokenizer('japanese'))
    
    # LexRankで要約を原文書の3割程度抽出
    summarizer = LexRankSummarizer()
    summarizer.stop_words = [' ']
    
    # 文書の重要なポイントは2割から3割といわれている?ので、それを参考にsentences_countを設定する。
    N = 3

    summary = summarizer(document=parser.document, sentences_count = N if len(corpus) < 100 else int(len(corpus)/100))
    #summary = summarizer(document=parser.document, sentences_count=1)
    
    str = ''
    for sentence in summary:
        str += (text[corpus.index(sentence.__str__())])
    return str

def main():
    fn1 = ''
    fn2 = ''
    pid = os.getpid()
    env = str(pid) + '$XXXX'
    #print(sys.argv)
    for i in range(len(sys.argv)):
        if i==1: fn1 = sys.argv[i]
        if i==2: fn2 = sys.argv[i]
    text = rd(fn1)
    rst= fn_start_document_summarize(text)
    print(u'文書要約結果:', rst)
    if fn2 == '':
        fn2 = env + '.txt'
    wf(fn2, rst.strip())
    return rst

if __name__ == '__main__':
    main()
   