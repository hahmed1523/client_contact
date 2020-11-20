'''This is a program to automate cleaning the client contact data'''
import os, pandas as pd, numpy as np, xlsxwriter as xl, pyautogui as p


def stat_rank(status):
    '''Return the rank for the status. Return 6 if its not in statusrank'''
    statusrank = {'Completed':1,'Pending':2, 'Worklisted':2, 'Custody Ended Partial Month':3, 'Abridged':4, 'Missed':5}
    if status in statusrank:
        return statusrank[status]
    else:
        return 6

def c_rank(case):
    '''Return the rank of the case type'''
    #Add a column to rank the Case Types 1-5
    caserank = {'Family Investigation':5, 'Treatment':4, 'Permanency':3, 'Guardianship':2, 'Adoption':1}
    if case in caserank:
        return caserank[case]
    else:
        return 6

#Get path for the csv files.
csvfilespath = p.prompt('Enter the full path to the client contact monthly csv files:', title ='File Location')

#Make a list of csv file names.
files = []
for file in os.listdir(csvfilespath):
    if file.endswith('.csv'):
        files.append(file)

#Create dataframe for each file
allrankdf = pd.DataFrame()
for file in files:
    df = pd.read_csv(os.path.join(csvfilespath,file))

    #Remove age group of '>=18'
    age_group_cond = df['Age Group'] != '>=18'
    df1 = df[age_group_cond].copy()

    #Make sure if 'Contact Complete?' column is 0 then 'Placement Setting?' is 0 as well.
    contact_complete_cond = df1['Contact Complete?']==0
    df1.loc[contact_complete_cond, 'Placement Setting?'] = 0

    #Check to see if there are any blanks in the 'Contact Due Date' column and if so then change 'Contact Complete?' and 'Placement Setting?' columns to 0.
    if np.where(df1['Contact Due Date'].isnull()):
        df1.loc[df['Contact Due Date'].isnull(), ['Contact Complete?','Placement Setting?']] = 0

    #Add a column to rank the Contact Status and Case Type
    df1.insert(10,'Status Rank', df1.loc[:,'Contact Status'].apply(stat_rank),True)
    df1.insert(6,'Case Rank', df1.loc[:,'Case Type'].apply(c_rank),True)

    #Sort by Status Rank, Placement Setting?, and Case Rank. Remove duplicates.
    df1 = df1.sort_values(by=['Client PID','Status Rank', 'Placement Setting?', 'Case Rank'], ascending = [True,True,False,True]).copy()

    #Remove duplicates
    df1 = df1.drop_duplicates(subset='Client PID', keep='first').copy()

    #Drop Status Rank and  Case Rank columns
    df1 = df1.drop(['Status Rank', 'Case Rank'], axis = 1)

    #Change data type for Schedule Month into datetime
    df1.loc[:,'Schedule Month'] = pd.to_datetime(df1.loc[:,'Schedule Month'])
    
    #Change any Pending Contact Status to 0 for Contact Complete? and Placement Setting?
    if np.where(df1['Contact Status']=='Pending'):
        df1.loc[df1['Contact Status']=='Pending', ['Contact Complete?', 'Placement Setting?']] = 0

    #Append all DataFrames into one Dataframe
    allrankdf = allrankdf.append(df1, ignore_index=True).copy()

#Sort by Schedule Month
allrankdf = allrankdf.sort_values(by='Schedule Month').copy()

#Adjust columns in excel and then Export to excel)
filename = 'client_contact_allranked.xlsx'
path = os.path.join(csvfilespath,filename)
writer = pd.ExcelWriter(path, engine = 'xlsxwriter', datetime_format = 'mm/dd/yyyy')
allrankdf.to_excel(writer, sheet_name = 'All Ranked', index=False)

workbook = writer.book
sheet = writer.sheets['All Ranked']

for col in range(18):               #Adjust the first 18 columns to size 15
    sheet.set_column(col,col, 15)

writer.save()

