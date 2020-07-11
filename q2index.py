import os
import split_folders
import re
from nltk.stem import PorterStemmer

def listofwords():
    lemmatizing=PorterStemmer ()
    fo=open("StopwordList.txt")
    x=fo.read()
    stoplistt=x.split()
    stoplistt=list (map(lambda x:x.lower(),stoplistt))
    fo.close()
    filelist=[] 
    dict={}
    d2={}
    path = os.getcwd()
    path+="\\bbcsport"
    folders=os.listdir(path)
    qq=re.compile("[^a-zA-Z]")
    
    for folder in folders:
        fname=path+"\\"+folder
        allfiles=os.listdir(fname)
        for f in allfiles:
            fname=path+"\\"+folder
            fname+="\\"+f
            lines=open(fname,"r")
            wordslist=lines.read()
            line = wordslist.rstrip()
            x=qq.split(line)
            x=list(filter(None,x))
            name=re.findall('[0-9]*',f)
            name=folder+"\\"+(name[0])
            for i in x:
                i=i.lower()
                if i in stoplistt:
                    continue
                i=lemmatizing.stem(i)
                # if len(i)<2:
                #     continue
                if i in stoplistt:
                    continue
                if i not in dict:
                    dict[i]=[]
                    dict[i].append(name)
                    d2[i]=1
                else:
                    if name not in dict[i]:
                        dict[i].append(name)
                        d2[i]=d2[i]+1
            lines.close()
            filelist.append(name)
            fname=""
    keytodelete=[]
    for k , v in d2.items():
        if (v)<3:
            keytodelete.append(k)
            
    for i in keytodelete:
        if i in dict:
            del dict[i]
        
    wordfile=open("vocabulary.txt","w")
    wordfile.write(str(dict))
    wordfile.close()
    
    lst=list(dict.keys())
    for i in range(len(lst)):
        lst[i]=lst[i].lower()
    lst.sort()  #returns a list of words in sorted order so that we can check the index of a particular word
    wordfile=open("listfwords.txt","w")
    wordfile.write(str(lst))
    wordfile.close()
    
    wordfile=open("filelist.txt","w")
    wordfile.write(str(filelist))
    wordfile.close()
    
    print("Total Terms : ",len(lst))

# listofwords()
