#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 31 19:27:27 2017
takes .xlsx file from Santander and correctly formats - giving us back .xlsx file to be put into spreadsheet
RAW DATA wants columns like so ["Date", "Transaction Type", "Merchant/ Description","Debit/Credit","Balance"]
@author: DewiGould
"""

import pandas as pd
from pandas import DataFrame
from collections import defaultdict
import math

df = pd.read_excel("RAW_DATA_FILE_TO_READ.xlsx",skiprows = 3)

#rename columns
df.columns = ["column 1", "Date", "column 2", "Description","column 3","Money In","Money Out","Balance","column 4"]

#delete empty columns
del df["column 1"], df["column 2"], df["column 3"], df["column 4"]

df.drop(0, inplace = True)

#Use words in description to ascribe a transaction type to each transaction
transaction_type = {"CARD": "Card","ATM": "Cash", "FASTER":"Bank", "DIRECT DEBIT": "Bank", "CASH":"Cash", "BANK": "Bank", "STANDING ORDER":"Bank", "BILL PAYMENT":"Bank", "CREDIT":"Bank"}
#remove un-important phrases from description
phrases_to_remove = ["CARD PAYMENT TO", "DIRECT DEBIT PAYMENT TO", "CASH WITHDRAWAL AT", "BILL PAYMENT VIA FASTER PAYMENT", "BILL PAYMENT TO", "FASTER PAYMENTS RECEIPT", "STANDING ORDER VIA FASTER PAYMENT", "CREDIT FROM"]

new_dataframe = defaultdict(list)

for index, row in df.iterrows():

    new_dataframe[index].append(row["Date"]) #add date
    
    description = row["Description"].encode('utf8').split(",")[0] #convert to string, only take info before comma
    
    for word in transaction_type.keys():
        if len(new_dataframe[index]) == 1:
            if word in description:
                new_dataframe[index].append(transaction_type[word])   #add transaction type
    if len(new_dataframe[index]) == 1:
        new_dataframe[index].append("Other")        #add other as transaction type if nothing found
 
    count = 0
    for phrase in phrases_to_remove:
        if phrase in description:
            new_dataframe[index].append(description.replace(phrase,""))      #remove unwanted phrases
            count +=1

    if count == 0:      #if no phrases found, inlucde un-edited version
        new_dataframe[index].append(description)
            
    if math.isnan(row["Money In"]) is True:
        new_dataframe[index].append(-1*row["Money Out"])  #add debit/credit
    else:
        new_dataframe[index].append(row["Money In"])
        
    new_dataframe[index].append(row["Balance"])
        


new_df = DataFrame.from_dict(new_dataframe,orient="index")

print new_df.head()


    
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('FORMATTED_DATA.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
new_df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.save()
