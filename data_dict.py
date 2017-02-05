


def data_dict(df, outfile):

    def get_top_vals(col):
        vals = col.value_counts()
        out = vals[0:min(len(vals),5)].to_dict()
        return str(sorted(out.items(), key=lambda kv: kv[1], reverse=True))

    def get_col_widths(dataframe):
        idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
        return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(str(col))]) for col in dataframe.columns]

    dictionary = pd.DataFrame(df.dtypes.astype('U'))
    dictionary.reset_index(inplace=True)
    dictionary.columns = ['Variable','Dtype']
    dictionary['Description'] = ''
    dictionary['Missing'] = dictionary['Variable'].apply(lambda x: sum(df[x].isnull()))
    dictionary['TopValues'] = dictionary.apply(lambda row: get_top_vals(df[row['Variable']]) if row['Dtype']=='object' else 'NA', axis=1)
    dictionary['Mean'] = dictionary.apply(lambda row: np.mean(df[row['Variable']]) if row['Dtype'] in ['float64','int64'] else 'NA', axis=1)
    dictionary['Min'] = dictionary.apply(lambda row: np.min(df[row['Variable']]) if row['Dtype'] in ['float64','int64'] else 'NA', axis=1)
    dictionary['Max'] = dictionary.apply(lambda row: np.max(df[row['Variable']]) if row['Dtype'] in ['float64','int64'] else 'NA', axis=1)
    dictionary['Mode'] = dictionary['Variable'].apply(lambda x: sme[x].value_counts(dropna=False).idxmax())
    dictionary['Mode%'] = dictionary['Variable'].apply(lambda x: sme[x].value_counts(dropna=False).max()/float(len(sme)))
    dictionary['Notes'] = ''

    writer = pd.ExcelWriter(outfile, engine='xlsxwriter')
    dictionary.to_excel(writer, sheet_name='Data Dictionary', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Data Dictionary']

    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format2 = workbook.add_format({'num_format': '0%'})
    format3 = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'left'})

    for i, width in enumerate(get_col_widths(df)):
        worksheet.set_column(i-1, i-1, min(width, 50))

    worksheet.set_row(0, 20, format3)

    writer.save()

    return dictionary