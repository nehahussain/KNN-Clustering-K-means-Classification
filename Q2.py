import os
import re
from nltk.stem import PorterStemmer
import q2index
import math
import ast
from scipy.spatial import distance
import pandas as pd
import random as rd
import numpy as np

q2index.listofwords()               #to create the vocabulary of the files
lemmatizing= PorterStemmer()
def returndict():                   #returns the vocabulary in the form of dict
    f=open("vocabulary.txt","r")
    r=f.read()
    dc=ast.literal_eval(r)
    f.close()
    return dc
#dictionary of the terms and its posting list, filelist of the training set, stopwordslist
Dictionary=returndict()

def returnfilelist():       #returns the total filelist 
    filee=open("filelist.txt","r")
    r=filee.read()
    r=ast.literal_eval(r)
    filee.close()
    return r

filelist=returnfilelist()

def returnlistofwords():    #returns the list of all words in the vocabulary
    fl=open("listfwords.txt","r")
    r=fl.read()
    r=ast.literal_eval(r)
    fl.close()
    return r

lstofword=returnlistofwords()

#returns a list of document frequency of terms
def get_df(): 
    templist=[]
    for i in range(len(lstofword)):
        templist.append(0)
    for i in lstofword:
        templist[lstofword.index(i)]=( len(Dictionary[i]))
    return templist

dflist=get_df()         #list of doc frequency of terms

N=len(filelist)        #total number of documents 

print("creating VSM")
def VSMofDoc():
    dict={}                      #dictionary for VSM
    temp=[]                     #temp list for document vector
    path = os.getcwd()          #returns the path of the current folder 
    path+="\\bbcsport"
    folders=os.listdir(path)        #to get all the folders in the bbc folder
    qq=re.compile("[^a-zA-Z]")
    path+="\\"
    for folder in folders:      #iterate through folders 
        fname=path+folder
        allfiles=os.listdir(fname)

        for f in allfiles:                     # for each file present in the folder
            temp=[]
            for jj in range(len(lstofword)):    # initilizing the vector with zero 
                temp.append(0)
            
            fnname=fname+"\\"+f                   #file name
            lines=open(fnname,"r")               #open a file 
            wordslist=lines.read()              #read the data from the file
            line = wordslist.rstrip()           #stripping the data
            x=qq.split(line)                    #splitting the data based on the regular expression
            x=list(filter(None,x))                  #deleting any none element present in the list because of split

            for i in x:                             #for every word in the list
                i=i.lower()                         #lowercase the term
                i=lemmatizing.stem(i)                 #stemming the word
                if i in lstofword:                      #if the term is present in the ist of words
                    temp[lstofword.index(i)]+=1            #incremening the term frequency
        
            name=re.findall('[0-9]*',f)
            dict[folder+"\\"+name[0]]=temp.copy()               #assigning the vector to its doc id
            lines.close()                                    #closing the file
            temp.clear()
        fname=""
    
    for ii,j in dict.items():
        for iii in range(len(j)):
            if j[iii]!=0:
                product=( j[iii] )*( math.log( N/dflist[iii] ,2) )              #multiplying the terms with its document frequency 
                j[iii]=product
    
    #write the VSM to a file  
    wordfile=open("DocumentVector.txt","w")
    wordfile.write(str(dict))
    wordfile.close()
    return dict
    # return
dictofDoc=VSMofDoc()

#read the VSM file
# f=open("DocumentVector.txt","r")
# r=f.read()
# dictofTrain=ast.literal_eval(r)
# f.close()
print("Vsm created")
K=5
rd.seed(10000)
Centroids=[]
Clusters={}             
forcentroid=rd.sample(filelist,K)       #randompy select the file names for initial centroid
# print(forcentroid)
for i in range(K):                      #initial centroids 
    Clusters[i]=[]
    temp = (dictofDoc[forcentroid[i]]).copy()           #getting the vector from dictionary of the file, which is randomly selected for the centroids
    Centroids.append(temp.copy())
    temp.clear()

difference=5
count=1
while difference!=0:                    #until the old centroid doesnt equal to new centroid
    difference=5
    for f in filelist:              #for every file present in the list
        dist=[]
        for jj in range(K):    # initilizing the vector with zero 
            dist.append(0)
        for i in range(K):
            numerator=np.dot(dictofDoc[f],Centroids[i])             #dot product of the file vector with the centroid
            magnitudeoffile=np.linalg.norm(dictofDoc[f])                #magnitude of the file vector
            magnitudeofcentroid=np.linalg.norm(Centroids[i])            #maagnitude of the centroid vector
            denominator=magnitudeoffile*magnitudeofcentroid
            Edistance=numerator/denominator                                 #calcullate the cosinesimilarity
            dist[i]=Edistance                                               #append the distance to list to check minimum dist and get the name of that minimum  dist file
            
        minimumdist=max(dist)
        indexofcluster=dist.index(minimumdist)
        Clusters[indexofcluster].append(f)              #append it to a cluster whse distance is the max as we calculated cosine similarity
    
    newCentroid=[]                                          #new centroid of the terms present in particular clusters
    for i in range(K):
        tempdict={}
        files_in_clus=Clusters[i]
        for i in range(len(files_in_clus)):
            temp2 = (dictofDoc[files_in_clus[i]]).copy()
            tempdict[files_in_clus[i]]=temp2.copy()
            temp2.clear()
        
        df=pd.DataFrame.from_dict(tempdict, orient="index")
        x=df.mean(axis=0)
        x=list(x)
        newCentroid.append(x.copy())
        tempdict.clear()
    
    for i in range(K):
        if newCentroid[i]==Centroids[i]:            #checking if new centroid is equal to old cenroid
            difference-=1
    if difference==0:
        break
    
    Clusters.clear()
    for i in range(K):
        Clusters[i]=[]
    Centroids.clear()           # clearing the old cluster
    
    Centroids=newCentroid.copy()      #assigning new centroids to centroids
    newCentroid.clear()
    print("Iteration : ",count)
    count+=1
    
    
Count={}
for i in range(len(Clusters)):
    Count[i]={
        "athletics":0,
        "cricket":0,
        "football":0,
        "rugby":0,
        "tennis":0
    }
    print("************* CLUSTER "+str(i)+"  **************")
    print(Clusters[i])
    print(len(Clusters[i]))
    for k in range (len(Clusters[i])):
        value=Clusters[i][k]
        find=re.findall("[a-zA-Z]*",value)
        Count[i][find[0]]+=1
    print("\n\n*******************************************\n\n")
purity=0
maxx=0
keyofmax=""
for k,v in Count.items():
    print("Cluster ",k,"  :    ")
    print(v)
    for ke, va in v.items():
        if maxx<va:
            maxx=va
            keyofmax=ke
    print("since ",keyofmax," has the max count so this cluster is of ", keyofmax," ")        
    purity+=maxx
    maxx=0
    print("\n") 

print("Purity : ",purity/737)   