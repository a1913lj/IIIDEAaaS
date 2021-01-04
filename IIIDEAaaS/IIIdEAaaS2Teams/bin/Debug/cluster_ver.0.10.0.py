'''

クラスタ分類

引数１：ファイル名
引数２：クラスタ数
引数３：解析種類　xml/txt ｘｍｌ形式/テキスト形式

'''

#https://qiita.com/y_itoh/items/158f004de6d243e3aeae
#https://qiita.com/ChenZheChina/items/42f1fcc763e88cb02cca
#https://qiita.com/zincjp/items/c61c441426b9482b5a48


#Doc2Vecとk-meansで教師なしテキスト分類
#https://qiita.com/shima_x/items/196e8d823412e45680e9
#

#from gensim import corpora
#import numpy as np
#from gensim.models.doc2vec import Doc2Vec, TaggedDocument
import sys
import codecs
from chardet.universaldetector import UniversalDetector
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from sklearn.decomposition import TruncatedSVD
from sklearn.preprocessing import Normalizer

import matplotlib.pyplot as plt
import xml.etree.ElementTree as ET


def detect_character_code(file):
    """
    pathnameから該当するファイルの文字コードを判別して、文字コードのdictを返す
 
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
    text = text.replace("？", "")
    text = text.replace("！", "")
    return text

def KM(X,n_clusters):
    km = KMeans(n_clusters=n_clusters,init='k-means++')
    km.fit_predict(X)
    return km.labels_

def seek(items):    
    vectorizer = TfidfVectorizer(use_idf=True)
    X = vectorizer.fit_transform(items)
    lsa = TruncatedSVD(int(len(items)/2))
    X = lsa.fit_transform(X)
    X = Normalizer(copy=False).fit_transform(X)
    return X
    
def sse(X):
    distortions = []

    for i  in range(1,11):                # 1~10クラスタまで一気に計算 
        km = KMeans(n_clusters=i,
                    init='k-means++',     # k-means++法によりクラスタ中心を選択
                    n_init=10,
                    max_iter=300,
                    random_state=0)
        km.fit(X)                         # クラスタリングの計算を実行
        distortions.append(km.inertia_)   # km.fitするとkm.inertia_が得られる
    
    plt.plot(range(1,11),distortions,marker='o')
    plt.xlabel('Number of clusters')
    plt.ylabel('Distortion')
    plt.show()
   
def parserTxt(file):
    '''
    1,Rectangle 3,アクティブラーニングツールとしてのアイデア発散・収束ソリューションの提案
    1,Rectangle 4,SMART Ideation
    1,Rectangle 5,アイデア一気通貫 虎の巻
    1,Rectangle 6,アイデア 収束 評価基準
    1,Rectangle 7,みんな諸葛孔明
    1,Rectangle 8,アイデア の実験
    1,Rectangle 9,逆アイデアから攻め
    1,Rectangle 10,ブレーンストーミングができるってことなのか？
    1,Rectangle 11,iTPS (Idea Total solution package)
    1,Rectangle 12,Idea Institute
    1,Rectangle 13,ideation for pro
    '''
    dic = {}
    
    with codecs.open(file, 'r', detect_character_code(file)) as f:
        contents = f.readlines()
    
    for c in contents:
        ss = c.strip().split(',')
        if len(ss) < 3:
            continue
        k=ss[0]
        Item = ss[1]
        Text = ''.join(ss[2:])
        if not k in dic.keys():
            dic[k] = []
            
        dic[k].append((Item, Text))
    #print (dic)
    return dic

def parserXml(fn):
    '''
    <root>
     <Slide>1
      <Rectangle>
       <Item>Rectangle 14</Item>
       <Text>付箋が足りないときは 下の付箋をキャンバス上に移動して 使ってください</Text>
      </Rectangle>
     </Slid>
    </root>
    '''
    tree = ET.parse(fn) 
    root = tree.getroot()
    
    dic = {}
    for slide in root.iter('Slide'):
        k = ''.join(slide.text.splitlines()).strip()
        for Rectangle in slide.iter('Rectangle'):
            Item = ''
            Text = ''
            for e in list(Rectangle):
                if e.tag == 'Item':
                    Item = ''.join(e.text.splitlines())
                elif e.tag == 'Text':
                    Text = clean(''.join(e.text.splitlines()))
                
                if len(Item) and len(Text):
                    #print(Item, Text)
                    if not k in dic.keys():
                        dic[k] = []
                    dic[k].append((Item, Text))
                    break;    
    #print (dic)            
    return dic                
   
def main():
    
    n_cluster = 5
    fn = ''
    ext = ''
    
    for i in range(len(sys.argv)):
        if i==1: fn=sys.argv[i]
        if i==2: 
            n_cluster = int(sys.argv[i])
            if n_cluster<=0: 
                n_cluster = 5
        if i==3: ext = sys.argv[i]
            
    print(fn, n_cluster, ext)
    
    if ext.lower() == 'txt':
        dic = parserTxt(fn)  
    else:
        dic = parserXml(fn)
        
    for k in dic:
        items = []
        Rectangles = []
        for v in dic[k]:
            Rectangles.append(v[0])
            items.append(v[1])
        if len(items) <=1: continue
        cluster = n_cluster
        if len(items) < 5:
            cluster = len(items)
        elif len(items) <= n_cluster:
            cluster = 5

        X = []
        X = seek(items)
        labels = KM(X,cluster)
        
        dic2 = {}
        i = 0
        for key in labels:
            if  key not in dic2:
                dic2[key] = []
            dic2[key].append((key, i, Rectangles[i], items[i])) 
            i +=1
        dic2 = sorted(dic2.items())
        dic[k] = dic2
    
    fn = ('.').join(fn.split('.')[:-1]) + '.rst'
    with open(fn, "w", encoding = "utf_8") as f:
        line = ','.join(('SlideNo.', 'ClusterNo.', 'IndexNo.', 'ShapeObjectType','Text'))
        f.write( line+'\n')
        for k in dic:
            for v in dic[k]:
                for v2 in v[1]:
                    line = ','.join((str(k), ','.join(map(str,v2))))
                    print(line)
                    f.write( line+'\n')
    return 

        
if __name__ == '__main__':
    main()
    


