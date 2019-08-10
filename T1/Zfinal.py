import os,sys,io
import xlrd
import xlsxwriter
import datetime
import time
import numpy as np
import pandas as pd
import  math

index =[]
column = []
# Function to calculate the Final P value
def PFun(z):
    p = 0.3275911
    a1 = 0.254829592
    a2 = -0.284496736
    a3 = 1.421413741
    a4 = -1.453152027
    a5 = 1.061405429
    if z < 0.0:
        sign = -1
    else:
        sign = 1
    x = math.fabs(z) / math.sqrt(2)
    t = 1.0 / (1.0 + p * x)
    erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * math.exp(-x * x)
    val=0.5 * (1.0 + sign * erf)
    print('P=',val)
    return val

# Paths to variace files
path = 'C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/input/'
path1 = 'C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/Spear Correlation/'
path2 = 'C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/Spear Correlation/FinalCorrelation.xlsx'
dest = 'C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/Final Output/Table1.xlsx'

# Storing all 54 user file names into two lists
for file in os.listdir(path):
    index.append(file)
    column.append(file)

#Creating an empty dataframe with NaN values and giving row and column names as User file names which we stored in lists
df = pd.DataFrame(data=np.nan,index=index, columns=column)

#Storing all correlation files into dataframes
df1 = pd.read_excel(os.path.join(path1,'w1w1.xlsx'))
df2 = pd.read_excel(os.path.join(path1,'w1w2.xlsx'))
df3 = pd.read_excel(os.path.join(path1,'w2w2.xlsx'))
df4 = pd.read_excel(os.path.join(path1,'w2w1.xlsx'))

#Appending all those dataframes into a single dataframe and setting up MuitiIndex
df5 = pd.concat([df1,df2,df3,df4],ignore_index=True)
df5 = df5.set_index(['User1', 'User2'])

#writer = pd.ExcelWriter(path1, engine='xlsxwriter')
#df5.to_excel(writer, sheet_name='Sheet1')
#writer.save()
#print(df5.loc[('w1t1ajb9b3.xlsx','w2t1ajb9b3.xlsx'),:])
#print(df5)

#Function to get Z value
def Zfun(a,b):
    r1 = df5.loc[('w1t1'+str(a),'w2t1'+str(a)),'P Value']
    r2 = df5.loc[('w1t1'+str(a),'w2t1'+str(b)),'P Value']
    r3 = df5.loc[('w2t1'+str(a),'w2t1'+str(b)),'P Value']
    z1= (1 / 2)*(math.log((1 + r1) / (1 - r1)))
    z2 = (1 / 2)*(math.log((1 + r2) / (1 - r2)))
    rm = ((r1*r1) + (r2*r2))/2
    f= (1- r3)/(2*(1-rm))
    h = (1-(f*rm))/(1-rm)
    N=16200
    N=N-3
    z = (z1-z2)*(math.sqrt(N))/(2*(1-r3)*h)
    print('Z=',z)
    return z

#Functions call for all possible spearman correlation value between different users
for i in index:
    for j in index:
       #print(i,j)
       z= Zfun(i,j)
       p=P = PFun(z)
       if math.isnan(p):
            p=0
       df.loc[i,j]=p

#Saving as an xlsx file which is the file table which will show case all correlatiosn between all the 54 users for 2 weeks
writer = pd.ExcelWriter(dest, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
