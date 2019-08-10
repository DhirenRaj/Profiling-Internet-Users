import os,sys
import xlrd
import xlsxwriter
import datetime
import time
import numpy as np
import pandas as pd
import scipy.stats
import math
c=''
pval =[]
rank1 = []
n1=[]
n2=[]
data1 = []
path1 = 'C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/W1T1/'
path2 = 'C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/W2T1/'

#This File gives us the spearman correlation for week2 & week1 for 10 secs window


for file1 in os.listdir(path2):
    #n1 = []
    d=0
    diff=[]
    rank1 =[]
    df1 = pd.read_excel(os.path.join('./W2T1',file1))
    # Storing all Usage(Doctet/Duration) column values of a file in a list
    dataset1 = df1['Usage']
    for file2 in os.listdir(path1):
        #n2 = []
        n1.append(file1)
        n2.append(file2)
        df2 = pd.read_excel(os.path.join('./W1T1', file2))
        # Storing all Usage(Doctet/Duration) column values of a file in a list
        dataset2 = df2['Usage']
        print(len(dataset2),len(dataset1))
        #manual calculation of Spearman correlation
        '''
        data1 =[]
        data2 =[]
        rank2 =[]
        #pval=[]
        df2 = pd.read_excel(os.path.join('./W1T1',file2))
        rank2 = df2['Usage'].rank(method='average')
        data1 = rank1-rank2
        data2 = data1*data1
        d = sum(data2)
        n=len(rank1)
        p = 1-((6*d)/(n*((n*n-1))))
        '''
        #Calculating Spearman Correlation with a built-in package function
        p = scipy.stats.spearmanr(dataset1, dataset2)[0]
        # Zeroing all the NaN values
        if math.isnan(p) == True:
            p = 0

        #Appending all the P values in a list
        pval.append(p)
        print(file1," ",file2," ",p)

#Storing all the lists into a dataframe and saving it as xlsx file
df3 = pd.DataFrame({"User1": n1, "User2": n2, "P Value": pval})
dest = 'C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/Spear Correlation/w2w1.xlsx'
writer = pd.ExcelWriter(dest, engine='xlsxwriter')
df3.to_excel(writer, sheet_name='Sheet1')
writer.save()
