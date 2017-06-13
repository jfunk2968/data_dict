import pandas as pd
import numpy as np

def data_dict(df, outfile):
    """Creates a nicely formatted Excel data dictionary for a pandas data frame.
    """
    def get_top_vals(col):
        vals = col.value_counts()
        out = vals[0:min(len(vals),5)].to_dict()  
        tops = []
        for i in sorted(out.items(), key=lambda kv: kv[1], reverse=True):
            s = str(i[0])+"  :  "+str(i[1])+"\n"
            tops.append(s)
        return "".join(tops)

    def get_col_widths(dataframe):
        return [max([len(str(s)) for s in dataframe[col].values] + [len(str(col))]) for col in dataframe.columns]

    def removeNonAscii(s): 
        return "".join(i for i in str(s) if ord(i)<128)

    dictionary = pd.DataFrame(df.dtypes.astype('U'))
    dictionary.reset_index(inplace=True)
    dictionary.columns = ['Variable','Dtype']
    dictionary['Description'] = ''
    dictionary['Missing'] = dictionary['Variable'].apply(lambda x: sum(df[x].isnull()))
    dictionary['TopValues'] = dictionary.apply(lambda row: get_top_vals(df[row['Variable']]) if row['Dtype']=='object' else 'NA', axis=1)
    dictionary['Mean'] = dictionary.apply(lambda row: np.mean(df[row['Variable']]) if row['Dtype'] in ['float64','int64'] else 'NA', axis=1)
    dictionary['Min'] = dictionary.apply(lambda row: np.min(df[row['Variable']]) if row['Dtype'] in ['float64','int64'] else 'NA', axis=1)
    dictionary['Max'] = dictionary.apply(lambda row: np.max(df[row['Variable']]) if row['Dtype'] in ['float64','int64'] else 'NA', axis=1)
    dictionary['Mode'] = dictionary['Variable'].apply(lambda x: df[x].value_counts(dropna=False).idxmax())
    dictionary['Mode%'] = dictionary['Variable'].apply(lambda x: df[x].value_counts(dropna=False).max()/float(len(df)))
    dictionary['Notes'] = ''
    
    dictionary['TopValues']= np.vectorize(removeNonAscii)(dictionary['TopValues'])

    writer = pd.ExcelWriter(outfile, engine='xlsxwriter')
    dictionary.to_excel(writer, sheet_name='Data Dictionary', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Data Dictionary']

    format1 = workbook.add_format({'num_format': '#,##0.00', 'align': 'right', 'valign': 'top', 'shrink': 'True'})
    format2 = workbook.add_format({'num_format': '0%', 'valign': 'top', 'shrink': 'True'})
    format3 = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'left', 'valign': 'top', 'shrink': 'True'})
    format4 = workbook.add_format({'text_wrap': 'True', 'align': 'right', 'valign': 'top', 'shrink': 'True'})

    worksheet.set_column('A:A', 20, format3)
    worksheet.set_column('D:D', 20, format1)
    worksheet.set_column('F:H', 20, format1)
    worksheet.set_column('J:J', 20, format1)
    worksheet.set_column('E:E', 20, format4)
    worksheet.set_column('I:I', 20, format4)
    worksheet.set_column('J:J', 20, format1)
    worksheet.set_column('C:C', 50)
    worksheet.set_column('K:K', 50)
    worksheet.freeze_panes(1, 1)

    writer.save()
    return dictionary

