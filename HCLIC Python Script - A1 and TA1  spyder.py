# -*- coding: utf-8 -*-
"""
Created on Thu May 30 16:35:56 2024

@author: freya.shrimpton
"""


import os
import pandas as pd
import re
import numpy as np

'1. DECIPHER THE CSV A// SHEET'
#Find out what sheets in the files have been made into csv
existing_csv = os.listdir('C:/Users/freya.shrimpton/Crisis/Data and Insights - Data Analysis and Reports/01 New Structure/08 Other Projects/24-05 HCLIC DATA/Dashboard Data')
sheets = []

for sheet in existing_csv:
    if len(sheet) <= 7:
        sheets.append(sheet.split('.csv')[0])
    
'2.START THE SHEETS LOOP - FOR EACH SHEET BEGIN THE PROCESS, DECIPHER THE FILES THAT NEED FORMATING '
for sheet in sheets:
    '1.1 Begin the process - locating the existing CSV and seeing what is in there already'
    #Read any csv that exists to see which sheets have already been existing sources
    csv = pd.read_csv('C:/Users/freya.shrimpton/Crisis/Data and Insights - Data Analysis and Reports/01 New Structure/08 Other Projects/24-05 HCLIC DATA/Dashboard Data/'+sheet+'.csv',index_col=0, header=0)
    existing_source = list(csv.source.drop_duplicates())
    #search the file for all of the workbooks
    files = os.listdir('C:/Users/freya.shrimpton/Crisis/Data and Insights - Data Analysis and Reports/01 New Structure/08 Other Projects/24-05 HCLIC DATA/data')
    #identify those that need to be added as thsoe that are not in the existing CSV
    files = [files for files in files if files not in existing_source]  #this should just be one file when updating
    
    '3. START THE FILES LOOP'
    for file in files:
        '3.1 Find the releavnt sheet - that may be different dut to naming conventions'
        xls = pd.ExcelFile('C:/Users/freya.shrimpton/Crisis/Data and Insights - Data Analysis and Reports/01 New Structure/08 Other Projects/24-05 HCLIC DATA/data/'+file) #list of all sheets in files
        file_sheets = xls.sheet_names
        file_sheets_clean = []
        for sheets1 in file_sheets:
            if len(sheets1) <= 4:
                file_sheets_clean.append(sheets1)
        file_sheet_df = pd.DataFrame({'sheets':file_sheets_clean})
        file_sheet = file_sheets_clean[int(np.where(file_sheet_df['sheets'].str.contains(sheet) == True)[0][0])]
        
        '3.2 Read the file with the correct file name'
        #sheet = 'A1'
        #file_sheet = 'A1'
        #file = 'Detailed_LA_202312.xlsx'
        data =  pd.read_excel(str('C:/Users/freya.shrimpton/Crisis/Data and Insights - Data Analysis and Reports/01 New Structure/08 Other Projects/24-05 HCLIC DATA/data/')+file,sheet_name = file_sheet)
        
        '3.3. creating quantitative fields that will get added at the end of the process'
        source_column = file
        sheet_column = (list(data.columns.values)[0].split('Table '))[1].split(' -')[0] #finds first table, or can use from list of sheets
        sheet_name_column = (list(data.columns.values)[0].split('- '))[1].split('England')[0].replace('\n','') #finds that first column and gets the data
        datemonth = file.split('_')[-1].split('.xlsx')[0]
        date_column = pd.to_datetime('01/'+datemonth[4:]+'/'+datemonth[0:4], format='%d/%m/%Y')
        #WORK OUT REST OF DATE COLUMNS IN POWERBI WHEN DONE
        
        '3.4. creating headers from the first few columns'
        
        #work out where first entry is to columns, if there is none then the colmn is empty and shold be dropped  
        first_col_index = data.iloc[:,0].first_valid_index() #first column shows where data starts
        columns = data.columns.values.tolist()
        headers = ['ONS','Area']
        for column in columns:
            first_index = data[column].first_valid_index()
            if (first_index == None) or first_index > first_col_index+1: #if its an empty column
              del data[column]
            elif first_index == first_col_index: #if its where the data starts
                pass
            elif  sheet == 'A1' : #going to have to make this sheet specific:
              headers.append(str(data[column][first_col_index-3]).replace('nan','')+' '+str(data[column][first_col_index-2]).replace('nan','')+' '+str(data[column][first_col_index-1]).replace('nan',''))
              
            elif data[column][first_col_index-2] is not None and str(data[column][first_col_index-1]) =='nan': #if cell above data is null, thne take the cell above
              headers.append(str(data[column][first_col_index-2]))
              
            elif str(data[column][first_col_index-2]) != 'nan' and data[column][first_col_index-1] is not None: #if cell abov edata is not null and neitehr is the one above, then take both
              headers.append(str(data[column][first_col_index-2]).split(' ')[0]+' '+str(data[column][first_col_index-1]).replace('nan',''))
              
            elif str(data[column][first_col_index-2]) == 'nan' and data[column][first_col_index-1] is not None: #if cell anove data is not null but cell above is null then get beginning of prev entry
              headers.append(str(headers[-1]).split(' ')[0]+' '+str(data[column][first_col_index-1]).replace('nan',''))
            else:
              pass
          
        '3.5 cleaning the headers so that they are all standardised'
        def remove(list):
         #   pattern = r'\W+'
            pattern = r'[^A-Za-z ]+'
            list = [re.sub(pattern, ' ', i) for i in list]
            return list
        
        headers = [h.replace('(000s)','Thousand') for h in headers] # clean headers     
        headers = [h.replace('Households','HHs') for h in headers] # clean headers
        headers = [h.replace('households','HHs') for h in headers] # clean headers
        headers = [h.replace('Number','No') for h in headers] # clean headers   
        headers = [h.replace('number','no') for h in headers] # clean headers   
        headers = [h.replace('Total','Tot') for h in headers] # clean headers        
        headers = [h.replace('\n',' ') for h in headers] # clean headers
        headers = remove(headers)  #remove numbers that confuse things 
        headers = [h.strip() for h in headers] # clean long spaces
        headers = [h.replace('   ',' ') for h in headers]
        headers = [h.replace('  ',' ') for h in headers]
        headers = [h.replace(' ','_') for h in headers]
        headers = [h.capitalize() for h in headers] #now they all have the correct headers
        data.columns = headers
        
        '3.6. filtering the data for null rows and trimming the beginning and the end'
        #find out row nmber where first value doe snot begin with E
        #data[data.columns[(data.values=='Notes').any(0)].tolist()] # where the notes bit is
        columns = data.columns.values.tolist()
        for column in columns: #for loop finds where the 'Notes' bit is and uses that to work out final column index
          if any(data[column].str.contains('Notes',na=False).tolist()) == True:
              last_col_index = int(np.where(data[column].str.contains('Notes') == True)[0])
              break
          else:
              pass
        data = data.iloc[first_col_index:last_col_index]
        data = data.drop(index=first_col_index+1)
        data = data.dropna(axis = 0, how = 'all').reset_index() # drop all fully null rows
        del data['index']
       

        
        '3.7 converting fields TO NUMERIC'
        #use isnumeric to identify numeric columns
        columns = data.columns.values.tolist()
        for column in columns:
            first_cell = str(data[column][0]).replace('.','') #remove dot for decimals
            if first_cell.isnumeric() == True:  #first instance of colmn is numeric
                data[column] = (pd.to_numeric(data[column], errors ='coerce').fillna(0))     
            else:
                pass
        
    
        '3.6. adding the columns first created in '
        data['sheet'] = sheet_column
        data['sheet_name'] = sheet_name_column
        data['date'] = date_column
        data['source'] = source_column
        
        '3.8. add new data tp the existing data - this is the final step of the  files loop'
        csv = csv._append(data)
      

          
        
        
    '3 - download new appended data for each sheet- THIS IS THE FINAL STEP IN THE SHEETS LOOP'
    csv.to_csv(str(sheet)+' - '+str(sheet_name_column)+'.csv') #for powerbi
    csv.to_csv(str(sheet)+'test.csv') #for next scriPt]
        
    
    
    
        
