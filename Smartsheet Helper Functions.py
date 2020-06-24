import pandas as pd
from pandas.io.json import json_normalize 
from datetime import datetime
import smartsheet
import pyodbc


tok ='<>'  # Your Smartsheet ODBC token
    
# Initialize client
smart = smartsheet.Smartsheet(tok)

#Function to get smartsheet as dataframe. Export sheet as csv and read as df. Faster than making the df cell by cell.
def get_sheet_as_df(smart , sheet_id):
    ss = smart.Sheets.get_sheet_as_csv(sheet_id , ".")  #dump csv into same directory as code. Specify own path in place of "." if needed.
    df = pd.read_csv(ss.filename)  #read that csv as dataframe
    return df   


#drop n rows from ss. numrows is the argument for n
def drop_rows(numrows):
    shee = smart.Sheets.get_sheet(sheet_id)
    sheet_dict = shee.to_dict()
    rowlist = [i['id']  for i in sheet_dict['rows']]  #list comprehension to get list of row ids of the smartsheet
    if (numrows > 100):  #If a large no of rows to be deleted
        for i in range(100 , numrows , 100):  #100 at a time
                smart.Sheets.delete_rows(sheet_id , rowlist[numrows-i:numrows-i+100])  #last 100 rows
        numrows = numrows-i  #to catch the balance above a multiple of 100. eg if numrows was 118, 18 are still pending
    if(numrows == 0):  #If numrows was a multiple of 100, then all have been processed
        return
    else:
        smart.Sheets.delete_rows(sheet_id , rowlist[-numrows:])  #delete the last n%100 rows using the Smartsheet sdk function.
    return

#append dataframe into smartsheet. This checks first if the SS row limit is exceeded. If it is, then it drops the extra rows first before appending.
def write_into_ss(df):    
    # Get all columns
    action = smart.Sheets.get_columns(sheet_id, include_all=True)
    columns = action.to_dict()  #JSON to dictionary   
    
    #Read the existing sheet as a df
    ssdf = get_sheet_as_df(smart , sheet_id)   #df of the smartsheet
    idlist = list(ssdf['id']) #List of all the document id's in the smartsheet till now
    if (ssdf.shape[0] + df.shape[0] >=19990): #if ss+incoming df is close to 20,000 rows, then we will need to make place for new additions. This is because SS has a limit of 20K rows.
        drop_rows(df.shape[0]) #drop ss rows equal to no of rows to be appended
    
    df = df.sort_values(by = ['<Your Column Name>'])  #Sort on name of the column you need to maintain sorted.
    #Nested loop to read each cell of the dataframe and copy that into smartsheet cell by cell
    for i in range(df.shape[0]):  
        if (df.iloc[i]['id'] in idlist):   #If the current row's docid is already existing in ss, then skip that row
            continue
        row = smartsheet.models.Row()   #initialize a new row variable for each row in the df
        row.to_top = True   #Append new rows into the top of the Smartsheet
        for j in range(df.shape[1]):  #Loop the columns for that row
            content = makecell( columns['data'][j]['id'] , df.iloc[i][j] )  #Pass the column id and dataframe cell value
            row.cells.append(content)
        smart.Sheets.add_rows(sheet_id,[row])
           
    return 0

#Create entry for a single cell
def makecell(col_id , value):
    cell = smartsheet.models.Cell()
    cell.column_id = col_id
    cell.value = str(value)   #For test purpose, make into string.
    return cell

