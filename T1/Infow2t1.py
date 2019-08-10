import os,sys
import xlrd
import xlsxwriter
import datetime
import time
import pandas as pd

# variable declarations
nj=0
slots=[]
slot=[]
rt=''
epoch = ''
doctet =''
dur=''
val =''
week_days=[11,12,13,14,15]
start_t = 0
end_t= 0
sum=0
count=0
#src = os.getcwd()
src = 'C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/input/'

for file in os.listdir(src):
    nj=0
    pstop=[]
    pstart=[]
    usage=[]
    print(file)
    filename = file
    #window size, here its 10 seconds
    t_slot = datetime.timedelta(0,10)
    #reading excel file into dataframe
    df= pd.read_excel('.\input/'+file)
    #dropping all values with zeros in Duration column and resetting its index
    df=df[df.Duration != 0]
    df=df.sort_values('Real First Packet')
    df=df.reset_index(drop=True)

    slots = []
    for i in week_days:
        # Getting packet usage values only the ist week Mon to Fri (4th to 8th)
        #In this loop we first create all the possible 10 sec time slots from 4th to 8th from 8AM to 5PM and store them in lists
        start_t = datetime.datetime(2013,2,i,8,0,0)
        end_t = datetime.datetime(2013,2,i,8,0,10)

        # Considering only from 8AM to 5PM
        while start_t.hour>=8 and end_t.hour!=18:
            slot=[start_t,end_t]
            slots.append(slot)
            start_t = start_t + t_slot
            end_t = end_t + t_slot
            if end_t.hour == 17:
                slot = [start_t, end_t]
                slots.append(slot)
                break
    #print(len(slots))
    for i in range(len(slots)):
        #Here we check for the avaliable packets in that time frame (10 secs)
        slot=slots[i]
        sum = 0
        count =0

        for j in range(nj,len(df.axes[0])):
            #epoch time is converted into local time
            rt = df.loc[j,'Real First Packet']/1000
            rt = datetime.datetime.fromtimestamp(rt)
            if  rt<slot[1]:
                if rt>=slot[0]:
                    # doctet/duration is calculated and stored in a list
                    dur = df.loc[j,'Duration']
                    doctet = df.loc[j,'doctets']
                    val = doctet/dur
                    sum=sum+val
                    count = count+1
            else:
                nj=j
                break
    #print(nj)
        if count>0:
            sum = sum/count
        else:
            sum=0

        #Apppending all the values into lists
        usage.append(sum)
        pstart.append(slot[0].strftime('%Y-%d-%m %H:%M:%S'))
        pstop.append(slot[1].strftime('%Y-%d-%m %H:%M:%S'))
        print(slot[0],' ',slot[1],' ',sum)
    #print(len(pstop),len(pstart),len(usage))
    # Storing all the lists as columns in a Dataframe and saving it as a xlsx time
    df1 = pd.DataFrame({"Start Time":pstart,"End Time":pstop,"Usage": usage})
    dest ='C:/Users/Dhiren Raj/Downloads/Information Security _ Privacy Material/W2T1/'+ 'w2t1' +filename
    writer = pd.ExcelWriter(dest, engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='Sheet1')
    writer.save()

#forusage=df['doctets']/df['Duration']
#x= datetime.datetime.fromtimestamp(epoch)



