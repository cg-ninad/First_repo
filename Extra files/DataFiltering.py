
import pandas as pd


def filter_countrywise_data(filename):
    try:
        countrywise_file = "SpeakUp/Country Wise Data.xlsx"
        df = pd.read_excel(filename,sheet_name="input")
        violaine_df = pd.read_excel(filename,sheet_name="Anne-Violaine")
        dict_of_country = dict(iter(df.groupby('Country')))
        writer = pd.ExcelWriter(countrywise_file, engine='xlsxwriter')
        for key, val in dict_of_country.items():
            val.to_excel(writer, sheet_name=key, header=True, index=False)

        violaine_df.to_excel(writer, sheet_name='Anne-Violaine', header=True, index=False)
        writer.save()
        return countrywise_file
    except Exception as err:
        print(err)



