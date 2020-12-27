# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import os
import pathlib
import sys
import pandas as pd
import codecs
from chardet.universaldetector import UniversalDetector

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


for a in sys.argv:
   fls = a 

#fldr = os.path.dirname(r'C:\Users\2ru\Desktop\PBL\Work\harman.csv')
#df = pd.read_csv(r'C:\Users\2ru\Desktop\PBL\Work\harman.csv',encoding='SHIFT JIS')

fldr = os.path.dirname(fls)
df = pd.read_csv(fls,detect_character_code(fls))

df=df.drop(df.columns[0], axis=1)

import matplotlib.pyplot as plt
#%matplotlib inline
import pandas as pd
#from pandas.tools import plotting
from pandas import plotting

#plotting.scatter_matrix(df[df.columns[1:]], figsize=(6,6), alpha=0.8, diagonal='kde')   #全体像を眺める
#grr = pd.plotting.scatter_matrix(df[df.columns[1:]], figsize=(6,6), alpha=0.8, diagonal='kde')
grr = pd.plotting.scatter_matrix(df[df.columns[2:]], figsize=(6,6), alpha=0.8, diagonal='kde')
#plt.show()
from sklearn.cluster import KMeans

kmeans_model = KMeans(n_clusters=len(df.columns)-1, random_state=10).fit(df.iloc[:, 1:])
labels = kmeans_model.labels_
#labels
a_i = [] 
for i in range(len(df)):
    a_i.append(i)
df['Cluster_ID']=labels
df.insert(0,'',a_i)
df.insert(0,'GroupNo',999)
#df['GroupNo']=999
df
#df['GroupNo']=999
#df
#df.query("Cluster_ID=='0'")
df_R = df.sample(frac=1)
#df_R.to_csv("k-means_before.csv", encoding="shift_jis", index=False)
CI_0=df_R[(df_R["Cluster_ID"] == 0)]
CI_1=df_R[(df_R["Cluster_ID"] == 1)]
CI_2=df_R[(df_R["Cluster_ID"] == 2)]
CI_3=df_R[(df_R["Cluster_ID"] == 3)]
CI_0.to_csv(fldr + "\\CI_0.csv", encoding="shift_jis", index=False)
CI_1.to_csv(fldr + "\\CI_1.csv", encoding="shift_jis", index=False)
CI_2.to_csv(fldr + "\\CI_2.csv", encoding="shift_jis", index=False)
CI_3.to_csv(fldr + "\\CI_3.csv", encoding="shift_jis", index=False)
