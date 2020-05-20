import ReadMail
import Quote
import pandas as pd
import PyPDF2
'''import config'''
from configparser import ConfigParser
configur = ConfigParser()
configur.read('Config.Ini')

#import trail
#import tabula
#import camelot
#import pdftotree
Output_df = pd.DataFrame(columns=['Supplier Name', 'QuoteID', 'Item Name','Quote Sent'])

Messages=ReadMail.get_Mail_Messages()
for Message in Messages:
    before_count=0
    after_count=0
    Out_Falg=False
    HTMl_Body=Message.HTMLBody.replace('\r','').replace('\n','')
    Str_Body=Message.body.replace('\r','')
    str_quote=Quote.get_quote(Str_Body)
    df_Quotes=Quote.get_df_quote(str_quote)
    Quote_Regex=Quote.get_quote_regex(df_Quotes)
    supplier=Quote.get_supplier(Str_Body)
    subFolderItemAttachments = Message.Attachments
    nbrOfAttachmentInMessage = subFolderItemAttachments.Count
    for attachment in subFolderItemAttachments:
        if '.xlsx' in str(attachment):
            attachment.SaveAsFile(configur.get('FilePath','OutputFolder') + '\\'+ attachment.FileName)
            xls = pd.ExcelFile(configur.get('FilePath','OutputFolder') + '\\'+ attachment.FileName)
            for sheet in xls.sheet_names:
                df = pd.read_excel(configur.get('FilePath','OutputFolder') + '\\'+ attachment.FileName, sheet_name=sheet, dtype=str)
                print(df)
                if Output_df.empty:
                    before_count=0
                else:
                    before_count=len(Output_df.index)
                Output_df=Quote.get_data_from_Excel_table(df,Quote_Regex,df_Quotes,Output_df,supplier,str_quote)
                if Output_df.empty:
                   after_count=0
                else:
                    after_count=len(Output_df.index)
                if after_count>before_count:
                    Out_Falg=True
        elif '.pdf' in str(attachment):
            attachment.SaveAsFile(configur.get('FilePath','OutputFolder')+ '\\'+ attachment.FileName)
            #'df=camelot.read_pdf('C:\\Users\\slice\\NLP POC\\NLP_Quotation\\OutPut' + '\\'+ attachment.FileName)
            pdf_file=configur.get('FilePath','OutputFolder') + '\\'+ attachment.FileName
            #pdfparser(pdf_file)
            #x=trail.extractPdfText(pdf_file)
            pdfFileObj = open(pdf_file, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            pageObj = pdfReader.getPage(0)
            Str_Body=pageObj.extractText()
            Str_Body=Str_Body.replace('\r','')
            print(pageObj.extractText().replace('\r',''))
            #df=pdftotree.parse(pdf_file, html_path=None, model_type=None, model_path=None, favor_figures=True, visualize=False)
            
    if Out_Falg==False:
        InBody=Quote.get_in_bound_MailBody(HTMl_Body)
        if "<table" in InBody:
            InBody=Quote.get_in_bound_MailBody(HTMl_Body)
            if Output_df.empty:
                before_count=0
            else:
                before_count=len(Output_df.index)
            Output_df=Quote.get_data_from_table(InBody,Quote_Regex,df_Quotes,Output_df,supplier,str_quote)
            if Output_df.empty:
                after_count=0
            else:
                after_count=len(Output_df.index)
            
            if after_count>before_count:
                Out_Falg=True
    if Out_Falg==False:
        InBody=Quote.get_in_bound_MailBody(Str_Body)
        Output_df=Quote.get_data_from_body(InBody,Quote_Regex,df_Quotes,Output_df,supplier,str_quote)
print(Output_df)
writer = pd.ExcelWriter(configur.get('FilePath','OutputFolder')+'\\output.xlsx')
Output_df.to_excel(writer)
writer.save()



#if __name__ == '__main__':
  #  pdfparser(sys.argv[1])    
    