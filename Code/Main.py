import ReadMail
import Quote
import pandas as pd

Messages=ReadMail.get_Mail_Messages()
for Message in Messages:
    HTMl_Body=Message.HTMLBody.replace('\r','').replace('\n','')
    Str_Body=Message.body.replace('\r','')
    str_quote=Quote.get_quote(Str_Body)
    df_Quotes=Quote.get_df_quote(str_quote)
    Quote_Regex=Quote.get_quote_regex(df_Quotes)
    supplier=Quote.get_supplier(Str_Body)
    Output_df = pd.DataFrame(columns=['Supplier Name', 'QuoteID', 'Item Name','Quote Sent'])
    if "<table" in HTMl_Body:
        InBody=Quote.get_in_bound_MailBody(HTMl_Body)
        Output_df=Quote.get_data_from_table(InBody,Quote_Regex,df_Quotes,Output_df,supplier,str_quote) 
    else:
        InBody=Quote.get_in_bound_MailBody(Str_Body)
        Output_df=Quote.get_data_from_body(InBody,Quote_Regex,df_Quotes,Output_df,supplier,str_quote)
    print(Output_df)
        
    