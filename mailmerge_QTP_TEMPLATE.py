from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pandas as pd # import pandas library to work with Excel files
from pandas import DataFrame, Series
import openpyxl

# IMPORT DATA FROM EXCEL SHEET, SET THE INDEX AS THE "Specific Test" COLUMN ITEMS. COLUMNS DO NOT HAVE TO BE SET MANUALLY. 
data = pd.read_excel(r'Automate_form.xlsx', index_col=0)

#print(data) # IF YOU WISH TO PRINT "data" DATAFRAME FIRST
# CONVERT "data" DATAFRAME TO A PYTHON DICTIONARY
#data_dict = data.to_dict()
#print(data_dict)

# ASSIGN VALUES TO ITEMS FROM DICTIONARY "data_dict" ## Could change Item column text to normal text instead of machine parsable
# Make this into a library??? 
#Ground_Survival_Low_Temperature=data.loc['Ground_Survival_Low_Temperature'].Value
# Short-Time_Operating_Low_Temperature=data.loc['Ground_Survival_Low_Temperature'].Value
#Operating_Low_Temperature =data.loc['Operating_Low_Temperature'].Value
#Ground_Survival_High_Temperature =data.loc['Ground_Survival_High_Temperature'].Value
#Short-Time_Operating_High_Temperature=data.loc['Short-Time_Operating_High_Temperature'].Value
#Operating_High_Temperature=data.loc['Operating_High_Temperature'].Value
#Altitude=data.loc['Altitude'].Value
#Decompression=data.loc['Decompression'].Value
#Overpressure=data.loc['Overpressure'].Value
# FINSH EXCEL SHEET AND CODE TO IMPORT DATA INTO DATAFRAME

#Pupulate Word document mailmerge items with Python dictionary

# Save the TEMPLATE document in "template" variable
template = "QTP-TEMPLATE_mailmerge.docx"
# Save the template as a MailMerge object. Note: describe this better, not sure if accurately described
document = MailMerge(template)

# Print out to console, the get_merge_fields 
#print("Your mail merge fields are: ")
#print(document.get_merge_fields())

print("You're merge fields are:\n")
for x in document.get_merge_fields():
    print(x)
print("\n")

# Populates merge fields from python library. 
document.merge(
    Part_Number=str(data.loc['Part_Number'].Value),
    Product_Description=str(data.loc['Product_Description'].Value),
#    DO160_S4_Cat='A4',
#    DO160_S5_Cat='C', # Temperature Variation
#    DO160_S6_Cat='A', # Humidity
#    DO160_S7_Cat='B', # Operational Shocks and Crash Safety
#    DO160_S8_Cat='S', # Vibration ---should specify curve to be used, usually Curve B
    Ground_Survival_High_Temperature=str(data.loc['Ground_Survival_High_Temperature'].Value),
    ShortTime_Operating_High_Temperature='50',
    Ground_Survival_Low_Temperature=str(data.loc['Ground_Survival_Low_Temperature'].Value),
    ShortTime_Operating_Low_Temperature=str(data.loc['Short-Time_Operating_Low_Temperature'].Value),
    Operating_High_Temperature=str(data.loc['Operating_High_Temperature'].Value),
    Operating_Low_Temperature=str(data.loc['Operating_Low_Temperature'].Value),
    Altitude=str(data.loc['Altitude'].Value),
    Decompression=str(data.loc['Decompression'].Value),
    Overpressure=str(data.loc['Overpressure'].Value)
)


# Once script is ready for exporting Word document with mailmerge
document.write('text-output.docx')

# Later perhaps import data from an SQL Server db, MySQL db, or other.