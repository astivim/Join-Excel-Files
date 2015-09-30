import pandas as pd
import textwrap
# 
"""
Exploratory data analysis for DFD/NFN : http://www.doctorsfordoctors.ca/
Join data from several sources to one Excel sheet. 
Files are excel and csv files from UN, WHO and WB database about Nicaragua. 
Load every data file to separate dataframe. 
Every data frame will have the same column names.
Column names are years.
Concatenate all and save to one .xslx file.
"""

def insert_newline(string):
    """ For string whose length >= 30 characters
        insert '\n'. New line '\n' is inserted 
        approximately every 30 characters, at the end of the 
        closest word.  """
    slen = 30
    str_split = textwrap.wrap(string,slen)
    str_split = [ st + '\n' for st in str_split]
    s = ' '
    return s.join(str_split)
    
    
# ==WORLD HEALTH ORGANIZATION DATA==============================================
# WHO data --------------------------------------------------- 
# http://apps.who.int/gho/data/node.country.country-NIC?lang=en

DF_WHO =  pd.read_csv('WHO_Nicaragua.csv',header = 1)
who_cols_1 = DF_WHO.columns 
who_cols_2 = [w.strip() for w in who_cols_1]
for i in range(0,len(who_cols_1)):
    DF_WHO = DF_WHO.rename(columns ={who_cols_1[i]:who_cols_2[i]})

DF_WHO = DF_WHO.set_index(['Indicator'])

#Add  information on the Data Source; 
DF_WHO.insert(0,'Data source',"WHO")
 
#=====UNITED NATIONS DATA=======================================================
# http://esa.un.org/unpd/wpp/DVD/
# This file contains population information for all the countries, extract 
# info only for Nicaragua
WPP = pd.ExcelFile('WPP2015_POP_F01_1_TOTAL_POPULATION_BOTH_SEXES.XLS',
                    skiprows = 16)
Sheet_Names = WPP.sheet_names # get the sheet names
del Sheet_Names[-1] # last sheet are text notes so erase it

Data_List = [(WPP.parse(sheet,header = 16,index_col = 2)).
             loc['Nicaragua'] for sheet in Sheet_Names]  # get all the sheets to a list   
DF_WPP = pd.concat(Data_List, join = 'outer',axis=1) # and then merge to one data frame
DF_WPP = DF_WPP.transpose()
DF_WPP = DF_WPP.set_index(['Variant']) 
len_DF_WPP = len(DF_WPP)

#Add a column with information on the Data Source
DF_WPP.insert(0,'Data source',"UN")

# Delete columns we don't need
del DF_WPP['Index'] 
del DF_WPP['Country code']
del DF_WPP['Notes']

#===WORLD BANK ORGANIZATION DATA ===============================================
#http://data.worldbank.org/country/nicaragua
DF_WB = pd.read_csv('WORLD_BANK_NIC.csv',skiprows = 4,index_col = 2)
DF_WB.insert(0,'Data source',"WB")
del DF_WB['Unnamed: 59']
del DF_WB['Country Name']
del DF_WB['Country Code']
del DF_WB['Indicator Code']


# ===JOIN ALL DATA============================================================== 
DF_ALL = pd.concat([DF_WHO,DF_WPP, DF_WB],join = 'outer',axis = 0)

# Joining columns automatically also sortss the columns putting the 
# 'Data source column' in the end. Since all three data frames have the same
# column names, we can reorder columns simply by 
#DF_ALL = DF_ALL.reindex_axis(DF_WPP.columns, axis=1)

# Move index to column 0. Format of the index column cannot be specified. 
# xlsxwriter will be used to write the file to Excel format.
DF_ALL.insert(0,'INDICATOR',DF_ALL.index)

# Indicator column is a string describing variable that was measured. In some
# cases it is very long and the Excel cell length is very long. Strings longer
# than 30 characters are broken to fit in several lines rather than in only 
# one line.
DF_ALL['INDICATOR'] = DF_ALL['INDICATOR'].apply(insert_newline) 


# ===SAVE TO EXCEL FILE=========================================================
writer = pd.ExcelWriter('NICARAGUA_DATA.xlsx', engine='xlsxwriter')
DF_ALL.to_excel(writer, sheet_name='Nicaragua', index = False)

# Editing the Excel file
workbook = writer.book
worksheet = writer.sheets['Nicaragua']
worksheet.set_tab_color('blue')

indicator_cell_format = workbook.add_format({'text_wrap': True})
worksheet.set_column('A:A',35,indicator_cell_format)
worksheet.set_column('EW:EW',30,indicator_cell_format)

row_color_1 = workbook.add_format({'bg_color': '#A5A6BA'})
row_color_2 = workbook.add_format({'bg_color':'#EFF0F5'})

len_df_all = len(DF_ALL)
for row in range(0,len_df_all,2):
    worksheet.set_row(row, None, row_color_1)
    worksheet.set_row(row+1,None,row_color_2)

writer.save()
