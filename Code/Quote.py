# -*- coding: utf-8 -*-
"""
Created on Sun Feb 23 20:24:11 2020

@author: Shailendra
"""
import re
import pandas as pd
from bs4 import BeautifulSoup
import spacy
from spacy.pipeline import SentenceSegmenter#
nlp = spacy.load('en_core_web_sm')

def split_on_newlines(doc):
    start = 0 
    seen_newline = False
    for word in doc:
        if seen_newline:
            yield doc[start:word.i]
            start = word.i
            seen_newline = False
        elif word.text.startswith('\n'): # handles multiple occurrences
            seen_newline = True
    yield doc[start:]      # handles the last group of tokens

sbd = SentenceSegmenter(nlp.vocab, strategy=split_on_newlines)
nlp.add_pipe(sbd)

def get_quote(str_body):
    
    try:
        x=re.search(r'\nQuote NO: #[\d]{4}\n',str_body.rsplit('From: Shailendra Varma Sagi <shailu.codepixer@outlook.com')[1]).group()
    except IndexError:
        return('Invaild body');
    else:
         return(re.search(r'[\d]{4}',x).group())
         
def get_df_quote(Quote):
    
        df_Quotes = pd.read_excel('C:\\Users\\slice\\NLP POC\\NLP_Quotation\\Input\\Quotes.xlsx')
        df_Quotes = df_Quotes[ df_Quotes['Quote Sent'] == int(Quote) ]
        return(df_Quotes)

def get_quote_regex(df_Quotes):
    
    #if len(str_body.rsplit('Shailendra Varma Sagi <shailu.codepixer@outlook.com>'))==2:
     #   return('Invaild body');
    #else:
    QuoteRegex=''
    for index,row in df_Quotes.iterrows():
        if QuoteRegex=='':
            QuoteRegex=row['Description']
        else:
            QuoteRegex=QuoteRegex +'|'+ row['Description']
    return(QuoteRegex)    
    
def get_supplier(str_body):
    x=re.search(r'\nSupplier Name: [\w{8,16}]*\n',str_body.rsplit('From: Shailendra Varma Sagi <shailu.codepixer@outlook.com')[1]).group()
    return(x.replace('\nSupplier Name:','').replace('\n',''))
    
def get_in_bound_MailBody(body):
    return(body.rsplit('From: Shailendra Varma Sagi <shailu.codepixer@outlook.com')[0])
    
def get_item_col_no(df,Quote_regex):
   
    print(Quote_regex)
    item_col_no=-1
    for index, row in df.iterrows():
        c=0
        for col in row:
            print(col)
            if str(col) in Quote_regex:
                print(c)
                item_col_no=c
                return(item_col_no);
                break;
            c=c+1
    return(item_col_no);
    
def get_Currency_of_DataFrame(df):
    
    for index, row in df.iterrows():
        for col in row:
            if get_currency(str(col))!=0:
                return(get_currency(str(col)));
                break;
    return(0);
    
def get_cost_Col_Form_header_Table(df):
    c=0
    cost_col_no=-1
    for col in df.columns: 
        #print(col) 
        if is_cost(str(col)):
            cost_col_no=c
            break;
        c=c+1
    return(cost_col_no)
    
def get_cost_Col_Form_header(df):
    c=0
    cost_col_no=-1
    for col in df.iloc[0]:
        #print(col)
        if is_cost(str(col)):
            cost_col_no=c
            break;
        c=c+1
    return(cost_col_no)
    
    
def get_Cost_Col_From_row(df):
   
    cost_col_no=-1
    for index, row in df.iterrows():
        c=0
        for col in row:
            if get_currency(str(col))!=0:
                cost_col_no=c
                return(cost_col_no);
                break;
            c=c+1
    return(cost_col_no);

def get_data_from_Excel_table(df,Quote_regex,df_Quotes,Output_df,supplier,Quote):
    print(Quote_regex)
    Count=1
    item_col_no=get_item_col_no(df,Quote_regex)
    if item_col_no!=-1:
        #cost_col_no=get_cost_Col_Form_header_Table(df)
           #When code headers are given
        if get_cost_Col_Form_header_Table(df)!=-1:
            cost_col_no=get_cost_Col_Form_header_Table(df)
            per_item_col=is_per(str(df.columns[cost_col_no]))
            Currency=get_currency(str(df.columns[cost_col_no]))
                #Currency=get_currency(str(list(sub_df.columns.values)[cost_col_no]))
        elif get_Cost_Col_From_row(df)!=-1:
            cost_col_no=get_Cost_Col_From_row(df)
            per_item_col=False
        else:
            print('unable to find the cost col')
        if Currency==0:
            Currency=get_Currency_of_DataFrame(df)
        if Currency==0:
            Currency='$'
        if cost_col_no!=-1:
            for index, row in df.iterrows():
                Received_Quote=False
                if str(row[item_col_no])!='nan'and str(row[cost_col_no])!='nan' :
                    if len(re.findall(Quote_regex,row[item_col_no]))==1:
                        item=re.findall(Quote_regex,row[item_col_no])[0]
                        per_item_amt=is_per(row[cost_col_no])
                        if per_item_col or per_item_amt:
                            Count=int(df_Quotes[ df_Quotes['Description'] == item ]['Count'].iloc[0])
                             
                        if len(re.findall(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',row[cost_col_no]))==1:
                            print(row[cost_col_no])
                            print(re.search(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',row[cost_col_no]).group())
                            
                            Quote_sent=str(round((float(re.search(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',row[cost_col_no]).group().replace(',',''))*int(Count)),2))
                            if Currency==0 :        
                                if get_currency(row[cost_col_no]) != 0:
                                    Quote_sent=get_currency(row[cost_col_no])+Quote_sent
                                    Received_Quote=True
                                else:
                                        print('Unidentified currency please add it to the data base')
                            else:
                                Quote_sent=Currency+Quote_sent
                                Received_Quote=True
                if Received_Quote==True:                
                    new_row = {'Supplier Name':supplier, 'QuoteID':Quote, 'Item Name':item,'Quote Sent':Quote_sent}
                    Output_df=Output_df.append(new_row,ignore_index=True)
        
    return(Output_df)           
    
def get_data_from_table(Html_body,Quote_regex,df_Quotes,Output_df,supplier,Quote):
    print(Quote_regex)
    soup = BeautifulSoup(Html_body, "html.parser")
    for body in soup("tbody"):
        body.unwrap()
    try:    
        df = pd.read_html(str(soup), flavor="bs4")
    except ValueError:
        print('No Table found')
        return(Output_df);
    

    for sub_df in df:
        Count=1
        item_col_no=get_item_col_no(sub_df,Quote_regex)
        if item_col_no!=-1:
            cost_col_no=get_cost_Col_Form_header(sub_df)
            #When code headers are given
            if get_cost_Col_Form_header(sub_df)!=-1:
                cost_col_no=get_cost_Col_Form_header(sub_df)
                per_item_col=is_per(str(sub_df.iloc[0][cost_col_no]))
                #Currency=get_currency(str(list(sub_df.columns.values)[cost_col_no]))
            elif get_Cost_Col_From_row(sub_df)!=-1:
                cost_col_no=get_Cost_Col_From_row(sub_df)
                per_item_col=False
            else:
                print('unable to find the cost col')
            Currency=get_Currency_of_DataFrame(sub_df)
            if cost_col_no!=-1:
                for index, row in sub_df.iterrows():
                    Received_Quote=False
                    if str(row[item_col_no])!='nan'and str(row[cost_col_no])!='nan' :
                        if len(re.findall(Quote_regex,row[item_col_no]))==1:
                            item=re.findall(Quote_regex,row[item_col_no])[0]
                            per_item_amt=is_per(row[cost_col_no])
                            if per_item_col or per_item_amt:
                                 Count=int(df_Quotes[ df_Quotes['Description'] == item ]['Count'].iloc[0])
                                 
                            if len(re.findall(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',row[cost_col_no]))==1:
                                print(row[cost_col_no])
                                print(re.search(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',row[cost_col_no]).group())
                                
                                Quote_sent=str(round((float(re.search(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',row[cost_col_no]).group().replace(',',''))*int(Count)),2))
                                if Currency==0 :        
                                    if get_currency(row[cost_col_no]) != 0:
                                            Quote_sent=get_currency(row[cost_col_no])+Quote_sent
                                            Received_Quote=True
                                    else:
                                        print('Unidentified currency please add it to the data base')
                                else:
                                    Quote_sent=Currency+Quote_sent
                                    Received_Quote=True
                        if Received_Quote==True:                
                            new_row = {'Supplier Name':supplier, 'QuoteID':Quote, 'Item Name':item,'Quote Sent':Quote_sent}
                            Output_df=Output_df.append(new_row,ignore_index=True)
        
    return(Output_df)       
        
def get_data_from_body(srt_body,QuoteRegex,df_Quotes,Output_df,supplier,Quote):
    doc=nlp(srt_body)
    for sent in doc.sents:
        Received_Quote=False
        print(sent.text)
        #Check if item is present in sentence
        if len(re.findall(QuoteRegex,sent.text))==0 or sent.text =='\n':
            
            print(sent.text)
            print(" No  Next Sentence Please\n")
        else:# if regex is there
            if len(re.findall(QuoteRegex,sent.text))==1:# if we find only one line item in the sentence
                if is_per(sent.text):# check if the item value given is cost per item or total cost
                    Count=int(df_Quotes[ df_Quotes['Description'] == re.search(QuoteRegex,sent.text).group() ]['Count'].iloc[0])
                else:
                    Count=1
                Money_list=[ent for ent in nlp(sent.text).ents if ent.label_ == 'MONEY']# do we have nay money entities?
                print(Money_list)
                if len(Money_list)>=1: #if we have more than one money entity
                    if len(Money_list)==1:# only one money entity
                    #logic for sentence with oniy one Quote regex and only one money entity started
                        for ent in nlp(sent.text).ents:
                            if ent.label_ == 'MONEY':
                                if len(re.findall(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',ent.text))==1:
                                    Quote_sent=str(round((float(re.search(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',ent.text).group().replace(',',''))*int(Count)),2))
                                    if get_currency(sent.text) != 0:
                                        Quote_sent=get_currency(sent.text)+Quote_sent
                                    else:
                                        print('Unidentified currency please add it to the data base')
                                    Received_Quote=True
                                else:
                                    print('Unable to find actual money in money entity')
                    #logic for sentence with oniy one Quote regex and only one money entity was completed
                    #logic for sentence with only one quote regex and more than one money enttity to find the actual  money entity started
                    elif len(re.findall(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',sent.text))>=1 and get_currency(sent.text)!=0:
                        Quote_sent=str(round((float(get_nearest_no(sent.text).replace(',',''))*int(Count)),2))
                        Quote_sent=get_currency(sent.text)+Quote_sent
                        Received_Quote=True
                        #logic c and more than one money enttity to find the actual  money entity completed
                    else:
                        print("Unable to get currency from the string please check")
                    
                    #logic for for sentence with only one quote regex and one number and currency start
        
                elif len(re.findall(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',sent.text))==1 and get_currency(sent.text)!=0:
                    Quote_sent=str(round(float(re.search(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',sent.text).group().replace(',',''))*int(Count)),2)
                    Quote_sent=get_currency(sent.text)+Quote_sent
                    Received_Quote=True
             #logic for for sentence with only one quote regex and one number and currency completed
             #logic for for sentence with only one quote regex and more than number and currency start
                elif len(re.findall(r'[1-9](\d+)*(,\d{1,})*([\.][\d]+)?',sent.text))>=1 and get_currency(sent.text)!=0:
                    Quote_sent=str(round((float(get_nearest_no(sent.text).replace(',',''))*int(Count)),2))
                    Quote_sent=get_currency(sent.text)+Quote_sent
                    Received_Quote=True
                else:
                    print('No money component present')
            #logic for for sentence with only one quote regex and more than number and currency completed
            #remove . id present at the end and add to the data frame
                        
                if Received_Quote==True:
                    if Quote_sent[-1]=='.':
                        Quote_sent=Quote_sent[0:-1]
                    new_row = {'Supplier Name':supplier, 'QuoteID':Quote, 'Item Name':re.findall(QuoteRegex,sent.text)[0],'Quote Sent':Quote_sent}
                    Output_df=Output_df.append(new_row,ignore_index=True)
                
                else:
                    print("Too many regex strings in the statement")
    return(Output_df)

    
    
    
def get_nearest_no(text):
    c=0
    for token in nlp(text):
        if get_currency(token.text)!=0:
            currency_token=c
            break;
        c=c+1
    c=0
    Nearest_No=''
    for token in nlp(text):
        if token.tag_=='CD':
            if Nearest_No=='':
                Nearest_No=token.text
                diff=abs(currency_token-c)
            else:
                if diff>abs(currency_token-c):
                    Nearest_No=token.text
                    diff=abs(currency_token-c)
        c=c+1
    return Nearest_No;
        
def is_per(text):
    Flag=0
    df_strings= pd.read_excel('C:\\Users\\slice\\NLP POC\\NLP_Quotation\\Input\\per.xlsx')
    for index, row in df_strings.iterrows():
        if str(row['String']).lower() in text.lower():
            Flag=1
            return True;
    if Flag==0:
        return False;
    
def get_currency(text):
    Flag=0
    if isNumber(text):
        return 0;
    else:
        df_Symbols = pd.read_excel('C:\\Users\\slice\\NLP POC\\NLP_Quotation\\Input\\Currency Symbols.xlsx')
        for index, row in df_Symbols.iterrows():
            if str(row['Text']).lower() in text.lower():
                Flag=1
                return str(row['Symbol']);
                break
        if Flag==0:
            return 0;

def is_cost(text):
    Flag=0
    if isNumber(text):
        return False;
    df_strings= pd.read_excel('C:\\Users\\slice\\NLP POC\\NLP_Quotation\\Input\\Cost.xlsx')
    for index, row in df_strings.iterrows():
        if str(row['String']).lower() in text.lower():
            Flag=1
            return True;
    if Flag==0:
        return False;


def isNumber(s) : 
      
    for i in range(len(s)) : 
        if s[i].isdigit() != True : 
            return False
  
    return True
    
    
        
        
        
    
