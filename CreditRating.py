import pandas as pd
import json
import re

df_credit_rating = pd.read_excel(r"C:\Users\heman\Desktop\MyWorkSpace\cre\Credit_Rating_Details.xlsx")
sf_rating_details = pd.DataFrame(df_credit_rating["Rating Details"])

def extract_keys_values(row):
    try:
        json_data = re.sub(r"[^\x00-\x7F]+", "", row.replace("'", '"'))
        json_data_parsed = json.loads(json_data)
        row_dict = {}
        for entry in json_data_parsed:
            for key, value in entry.items():
                if key in row_dict:
                    row_dict[key].append(value)
                else:
                    row_dict[key] = [value]
        return pd.Series(row_dict)
    except Exception as e:
        print(f"Error: {e}")
        return pd.Series({})

sf_rating_details = sf_rating_details['Rating Details'].apply(extract_keys_values)
sf_rating_details = sf_rating_details.applymap(lambda x: x[0] if isinstance(x, list) else x).fillna('')

print(sf_rating_details)
pd.DataFrame(sf_rating_details)
sf_rating_details.head(10)
sf_rating_details.isnull().sum()
sf_rating_details.info()
rows_with_date = sf_rating_details[sf_rating_details['date'] == "2023-07-03"]

pd.DataFrame(rows_with_date)
sf_rating_details.rename(columns={'orgNumber': 'CIN'}, inplace=True)

sf_rating_details = pd.DataFrame(sf_rating_details)
sf_rating_details.head()
columns = sf_rating_details.columns.tolist()
columns.remove('CIN')
columns.insert(0, 'CIN')
sf_rating_details = sf_rating_details[columns]
print(sf_rating_details)
sf_rating_details.head
sf_rating_details = pd.DataFrame(sf_rating_details)
sf_rating_details
df_screener = pd.read_excel(r"C:\Users\heman\Desktop\MyWorkSpace\cre\Screener_CIN_Mapping.xlsx")
df_screener
merged_data = pd.merge(sf_rating_details, df_screener, on="CIN" )
merged_data = merged_data.drop(["companyName"], axis=1)
merged_data.head(1)
merged_data.rename(columns={'date': 'Date of Rating',
                        'agency': "Rating Agency", 
                        'amount':"Amount (Mn)",
                        'rating':"Rating", 
                        "instrument":"Instrument",
                        "ratingGrade":"Rating Grade",
                        "ratingStatus":"Rating Status",
                        'Screener Link':"Link"}, inplace=True)

merged_data.head()
merged_data.to_excel('Credit_Rating_sheet.xlsx', index=False)

