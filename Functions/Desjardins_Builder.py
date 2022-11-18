# Import Libraries

import warnings
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
import re
from textblob import TextBlob

def main(df, ref_table):
    warnings.filterwarnings('ignore')

    today_date = date.today().strftime("%m.%d.%y")

    rename_dict = {
        'Full Name' : 'Customer', 
        'Account Number' : 'Account Number',
        'Timing' : 'Duration',
        'Channel Details' : 'Geo Target',
        'SEM LOB Exclusions ' : 'Lob Exclusion',
        'Province' : 'State',
    }

    df.rename(mapper = rename_dict, axis=1, inplace = True)
    df['Duration'] = 12
    df['Account'] = 'Desjardins Agents Network - COOP Program'

    source_df = df[['Customer', 
                    'Account Number', 
                    'Website',
                    'Geo Target',
                    'Account',
                    'Lob Exclusion',
                    'State',
                    'Phone Number',
                    'Duration',  
        ]]

    # Standardize the Phone Number
    source_df['Phone Number'] = source_df['Phone Number'].replace(regex=r'[^0-9.]', value='')
    source_df['Phone Number'] = '(' + source_df['Phone Number'].str[:3] + ') ' + source_df['Phone Number'].str[3:6] + '-' + source_df['Phone Number'].str[6:]

    source_df['State'] = source_df['State'].str.split('-', expand=True)[0].str.strip()

    can_province_abbrev = {
    'Alberta': 'AB',
    'British Columbia': 'BC',
    'Manitoba': 'MB',
    'New Brunswick': 'NB',
    'Newfoundland and Labrador': 'NF',
    'Northwest Territories': 'NT',
    'Nova Scotia': 'NS',
    'Nunavut': 'NU',
    'Ontario': 'ON',
    'Prince Edward Island': 'PE',
    'Quebec': 'QC',
    'Saskatchewan': 'SK',
    'Yukon': 'YT'
    }

    province_abb_list = list(can_province_abbrev.values())

    # Ensure that all State/Province names are 2 Digit Abbreviations.
    def Province_abb(x):
        if (len(x['State']) > 2) & (x['State'] not in province_abb_list):
            x['State'] = can_province_abbrev[x['State']]
        return x['State']

    source_df['State'].fillna('', inplace=True)

    source_df['State'] = source_df.apply(Province_abb, axis=1)

    # Location is long format of Province
    can_province_names = {
    'AB': 'Alberta',
    'BC': 'British Columbia',
    'MB': 'Manitoba',
    'NB': 'New Brunswick',
    'NF': 'Newfoundland and Labrador',
    'NS': 'Nova Scotia',
    'NT': 'Northwest Territories',
    'NU': 'Nunavut',
    'ON': 'Ontario',
    'PE': 'Prince Edward Island',
    'QC': 'Quebec',
    'SK': 'Saskatchewan',
    'YT': 'Yukon'
    }
    source_df['Location'] = source_df['State'].map(can_province_names)

    for i, row in source_df['Geo Target'].iteritems():
        
        # ZIP Code finder
        a = ''
        for j in re.findall(r"([A-Z]\d[A-Z])(?!([A-Z]|[a-z]))", row.upper()):
            a += ' | ' + j[0]
        
        # Radius finder
        _ = re.findall(r"([0-9]{1,3}) MILE", row.upper())
        if len(_) == 1:
            a += ' | ' + _[0] + ' Miles Radius'
            
        source_df['Geo Target'][i] = a[3:]


    ## Spell checking LOB Exclusions
    from textblob import TextBlob

    def spelling_checker(word):
        sentence = TextBlob(word.lower())
        result = sentence.correct()
        return str(result.title())

    def isNaN(string):
        return string != string

    def spell_check(x):
        res = x['Lob Exclusion']
        if isNaN(res):
            return res 
        res = res.split(', ')
        words = []
        
        for word in res:
            word = spelling_checker(word)
            if word == 'Ev':
                word = 'EV' 
            if word == 'Conde':
                word ='Condo'
            words.append(word)
            
        res = ', '.join(words)
        return res

    source_df['Lob Exclusion'] = source_df.apply(spell_check, axis=1)

        ## Snippet Upload --------------

    ref = pd.read_excel(ref_table, sheet_name='Structured Snippet Template') # read ref table
    Snippet_df = source_df.copy()


    def Snippet_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-0000-000000', row['Account Number'], regex=True, inplace=True)
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Snippet_df = Snippet_update(Snippet_df)

    # # path to save 
    # path = f'{main_path}\Output\Snippet Upload'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Snippet_df.to_excel(f'{path}\\New Agent Snippet Upload - {today_date}.xlsx', 
    #                     sheet_name='Structured Snippet Upload', 
    #                     index=False)


        ## Sitelink Upload --------------

    ref = pd.read_excel(ref_table, sheet_name='Sitelink Template') # read ref table
    Sitelink_df = source_df.copy()

    def Sitelink_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-0000-000000', row['Account Number'], regex=True, inplace=True)
            data_for.replace('https://www.URL.com/localagent', row['Website'], regex=True, inplace=True)
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Sitelink_df = Sitelink_update(Sitelink_df)

    # # path to save 
    # path = f'{main_path}\Output\Sitelink Upload'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Sitelink_df.to_excel(f'{path}\\New Agent Sitelink Upload - {today_date}.xlsx', 
    #                     sheet_name='Sitelink Upload', 
    #                     index=False)

        ## Call Upload --------------

    ref = pd.read_excel(ref_table, sheet_name='Call Template') # read ref table
    Call_df = source_df.copy()

    def Call_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-0000-000000', row['Account Number'], regex=True, inplace=True)
            data_for['Phone number'] = row['Phone Number']
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Call_df = Call_update(Call_df)

    # # path to save 
    # path = f'{main_path}\Output\Call Upload'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Call_df.to_excel(f'{path}\\New Agent Call Upload - {today_date}.xlsx', 
    #                     sheet_name='Call Upload', 
    #                     index=False)


        ### BULK UPLOAD -------------------
        ## 1) Campaign Type ---------------

    ref = pd.read_excel(ref_table, sheet_name='Campaign Template') # read ref table
    Campaign_df = source_df.copy()

    # Full List of Regions
    Excluded_region = 'CA-BC | CA-MB | CA-NB | CA-NF | CA-NS | CA-AB | CA-PE | CA-QC | CA-SK | CA-NT | CA-NU | CA-ON | CA-YT'

    # All Countries except 'CA'
    Excluded_country = 'AF|AL|AQ|DZ|AS|AD|AO|AG|AZ|AR|AU|AT|BS|BH|BD|AM|BB|BE|BM|BT|BO|BA|BW|BV|BR|BZ|IO|SB|VG|BN|BG|MM|BI|BY|KH|CM|US|CV|KY|CF|LK|TD|CL|CN|TW|CX|CC|CO|KM|YT|CG|CD|CK|CR|HR|CY|CZ|BJ|DK|DM|DO|EC|SV|GQ|ET|ER|EE|FO|FK|GS|FJ|FI|FR|GF|PF|TF|DJ|GA|GE|GM|PS|DE|GH|GI|KI|GR|GL|GD|GP|GU|GT|GN|GY|HT|HM|VA|HN|HK|HU|IS|IN|ID|IQ|IE|IL|IT|CI|JM|JP|KZ|JO|KE|KR|KW|KG|LA|LB|LS|LV|LR|LY|LI|LT|LU|MO|MG|MW|MY|MV|ML|MT|MQ|MR|MU|MX|MC|MN|MD|ME|MS|MA|MZ|OM|NA|NR|NP|NL|BQ|CW|AW|SX|BQ|NC|VU|NZ|NI|NE|NG|NU|NF|NO|MP|UM|FM|MH|PW|PK|PA|PG|PY|PE|PH|PN|PL|PT|GW|TL|PR|QA|RE|RO|RU|RW|SH|KN|AI|LC|PM|VC|SM|ST|SA|SN|RS|SC|SL|SG|SK|VN|SI|SO|ZA|ZW|ES|EH|SR|SJ|SZ|SE|CH|TJ|TH|TG|TK|TO|TT|AE|TN|TR|TM|TC|TV|UG|UA|MK|EG|GB|GG|JE|TZ|VI|BF|UY|UZ|VE|WF|WS|YE|ZM|XK|'

    def Campaign_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
                    
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-0000-000000', row['Account Number'], regex=True, inplace=True)             
            
            to_remove = row['State']
            if to_remove == 'BC':
                data_for['Excluded region'] = Excluded_region.str[8:]
            else :
                data_for['Excluded region'] = Excluded_region.replace(f' | CA-{to_remove}', '')

            data_for['Metro'] = row['Geo Target']
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        
        # Columns Manually written for all rows.
        res['Advertiser bid strategy'] = ''
        res['Campaign start date (YYYY-MM-DD)'] = date.today().strftime("%Y-%m-%d")        
        year_ahead = (date.today() + relativedelta(months=+row['Duration']))
        res['Campaign end date (YYYY-MM-DD)'] = year_ahead.replace(day=pd.Period(str(year_ahead)).days_in_month).strftime("%Y-%m-%d")
        
        res['Excluded country'] = Excluded_country
        
        res['Row Type'] = 'Campaign'
        res['Action'] = 'Create|Pause'
        res['Advertiser'] = ''
        res['Account'] = 'Desjardins Agents Network - COOP Program'
        res['Campaign daily budget'] = 5
        res['Campaign monthly budget'] = ''
        res['Campaign lifetime budget'] = ''
        res['Campaign budget type'] = ''
        res['Campaign target network'] = 'Search'
        res['Desktop bid adjustment %'] = ''
        res['Mobile bid adjustment %'] = ''
        res['Tablet bid adjustment %'] = ''
        res['Campaign language target'] = 'EN'
        res['Ad rotation'] = 'Rotate Indefinitely'
        res['Ad delivery method'] = 'Standard'
        res['Label'] = ''
        res['Label ID'] = ''
            
        return res

    Bulk_df_Campaign = Campaign_update(Campaign_df)

        ## 2) Ad Group Type ---------------

    ref = pd.read_excel(ref_table, sheet_name='Ad Group Template') # read ref table
    AdGroup_df = source_df.copy()


    # Check the main data and add all LOB Exclusions to a main list.
    def LOB(x):
        Lob_Exclusions = x['Lob Exclusion'].split(', ')

        Lob_Exclusions = ['EV' if KW == 'Electric Vehicle' else KW for KW in Lob_Exclusions]
        Lob_Exclusions = ['Farm_Ranch' if KW == 'Farm/Ranch' else KW for KW in Lob_Exclusions]
        
        if 'Mobile Home' in Lob_Exclusions:
            Lob_Exclusions.extend(['Mobile Home', 'Trailer', 'Motorhome', 'RV'])    
        
        return Lob_Exclusions
        
    AdGroup_df['Lob Exclusion'] = AdGroup_df.apply(LOB, axis=1)


    def AdGroup_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-0000-000000', row['Account Number'], regex=True, inplace=True)
            
            data_for['Action'] = 'Create'
            i = 0
            for row2 in data_for['Ad Group']:
                if row2.split(' - ')[1] in row['Lob Exclusion']:
    #                 print(row2.split(' - ')[1], row['Lob Exclusion'])
                    data_for['Action'][i] = 'Create|Pause'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        res['Row Type'] = 'Ad Group'
        res['Account'] = 'Desjardins Agents Network - COOP Program'
        res['Ad group search max CPC'] = 3
        
        return res


    Bulk_df_AdGroup = AdGroup_update(AdGroup_df)


        ## 3) Ad Type ---------------

    ref = pd.read_excel(ref_table, sheet_name='Ad Copy Template') # read ref table
    Ad_df = source_df.copy()

    Ad_df['Customer First Name'] = Ad_df['Customer'].str.split(' ',expand=True)[0]
    Ad_df['Lob Exclusion'] = Ad_df.apply(LOB, axis=1)

    def Ad_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()

    # Action and Ad Group Column Edit        
            data_for['Action'] = 'Create'
            i = 0
            for row2 in data_for['Ad Group']:
                if row2.split(' - ')[1] in row['Lob Exclusion']:
    #                 print(row2.split(' - ')[1], row['Lob Exclusion'])
    #                 data_for['Action'][i] = 'Create|Pause'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            
    # Ad Description and Ad Title Columns Lenght Edit
            for column in ['Description Line 1','Description Line 2', 'Description Line 3', 'Description Line 4']:
                i = 0
                for row3 in data_for[column]:
                    if len(row3) + len(row['Customer']) - agentname_len > 90:
                        data_for[column][i] = data_for[column][i].replace('[agentname]', row['Customer First Name'])
                    else: 
                        data_for[column][i] = data_for[column][i].replace('[agentname]', row['Customer'])
                    i += 1

            for column in ['Headline 1', 'Headline 2', 'Headline 3','Headline 4', 'Headline 5', 'Headline 6', 'Headline 7', 
                        'Headline 8', 'Headline 9', 'Headline 10', 'Headline 11', 'Headline 12','Headline 13', 'Headline 14', 'Headline 15']:
                i = 0
                for row3 in data_for[column]:
                    if len(row3) + len(row['Customer']) - agentname_len > 30:
                        data_for[column][i] = data_for[column][i].replace('[agentname]', row['Customer First Name'])
                    else: 
                        data_for[column][i] = data_for[column][i].replace('[agentname]', row['Customer'])
                    i += 1
        

            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-0000-000000', row['Account Number'], regex=True, inplace=True)
            data_for['Final URL'].replace('https://www.URL.com/localagent', row['Website'], regex=True, inplace=True)
            
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
            
        res['Row Type'] = 'Ad'
        res['Account'] = 'Desjardins Agents Network - COOP Program'
        return res

    agentname_len = len('[agentname]')

    Bulk_df_Ad = Ad_update(Ad_df).rename(columns={'Description Line 1':'Ad description line 1',
                                    'Description Line 2':'Ad description line 2',
                                    'Description Line 3':'Ad description line 3',
                                    'Description Line 4':'Ad description line 4',
                                    'Path 1':'Ad path field 1', 
                                    'Path 2':'Ad path field 2',
                                    'Headline 1':'Ad title',
                                    'Headline 2':'Ad title 2',
                                    'Headline 3':'Ad title 3',
                                    'Headline 4':'Ad title 4',
                                    'Headline 5':'Ad title 5',
                                    'Headline 6':'Ad title 6',
                                    'Headline 7':'Ad title 7',
                                    'Headline 8':'Ad title 8',
                                    'Headline 9':'Ad title 9',
                                    'Headline 10':'Ad title 10',
                                    'Headline 11':'Ad title 11',
                                    'Headline 12':'Ad title 12',
                                    'Headline 13':'Ad title 13',
                                    'Headline 14':'Ad title 14',
                                    'Headline 15':'Ad title 15',
                                    'Final URL':'Ad landing page'
                                    })


        ## 4) Keyword Type ---------------

    ref = pd.read_excel(ref_table, sheet_name='Keyword Template') # read ref table
    Keyword_df = source_df.copy()

    Keyword_df['Customer First Name'] = Keyword_df['Customer'].str.split(' ',expand=True)[0]
    Keyword_df['Lob Exclusion'] = Keyword_df.apply(LOB, axis=1)

    def Keyword_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            
    # Action and Ad Group Column Edit        
            data_for['Action'] = 'Create'
            i = 0
            for row2 in data_for['Ad Group']:
                if row2.split(' - ')[1] in row['Lob Exclusion']:
    #                 data_for['Action'][i] = 'Create|Pause'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            
            
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-0000-000000', row['Account Number'], regex=True, inplace=True)
            
            
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        res['Row Type'] = 'Keyword'
        res['Account'] = 'Desjardins Agents Network - COOP Program'
        res['Ad group search max CPC'] = 3
        
        return res

    Bulk_df_Keyword = Keyword_update(Keyword_df).rename(columns={'Max CPC':'Keyword Max CPC',
                                            'Criterion Type':'Keyword match type'})

        ## 5) Merge four Subtable to One -----------------

    columns = list(Bulk_df_Campaign.columns) + list(Bulk_df_AdGroup.columns) + list(Bulk_df_Ad.columns) + list(Bulk_df_Keyword.columns)
    columns = list(dict.fromkeys(columns))
    # columns

    columns = [
        'Row Type',
        'Action',
        'Advertiser',
        'Account',
        'Campaign',
        'Advertiser bid strategy',
        'Campaign start date (YYYY-MM-DD)',
        'Campaign end date (YYYY-MM-DD)',
        'Campaign daily budget',
        'Campaign monthly budget',
        'Campaign lifetime budget',
        'Campaign budget type',
        'Campaign target network',
        'Desktop bid adjustment %',
        'Mobile bid adjustment %',
        'Tablet bid adjustment %',
        'Campaign language target',
        'Country',
        'Region',
        'Metro',         
        'City',
        'Ad rotation',
        'Ad delivery method',
        'Ad Group',
        'Ad group language',
        'Ad group search max CPC',
        'Ad group content max CPC',
        'Ad group native max CPC',
        'Ad group start date (YYYY-MM-DD)',
        'Ad group end date (YYYY-MM-DD)',
        'Ad ID',
        'Ad name',
        'Ad Type',
        'Ad title',
        'Ad title 2',
        'Ad title 3',
        'Ad title 4',
        'Ad title 5',
        'Ad title 6',
        'Ad title 7',
        'Ad title 8',
        'Ad title 9',
        'Ad title 10',
        'Ad title 11',
        'Ad title 12',
        'Ad title 13',
        'Ad title 14',
        'Ad title 15',
        'Ad description line 1',
        'Ad description line 2',
        'Ad description line 3',
        'Ad description line 4',
        'Ad path field 1',
        'Ad path field 2',
        'Ad landing page',      
        'Ad display URL',
        'Mobile display URL',
        'Ad mobile landing page',
        'Ad device preference',
        'Ad Sponsored By',
        'Keyword ID',
        'Keyword',
        'Keyword match type',
        'Keyword Max CPC',
        'Keyword landing page',
        'Keyword min strat bid',
        'Keyword max strat bid',
        'Label',
        'Label ID',
        'Keyword Param 2',  
        'Keyword Param 3',
        'URL Template',              
        'Excluded region',
        'Excluded country' 
    ]


    Merged_df = pd.DataFrame(columns=columns)
    # Merged_df


    bulk_df = pd.concat([Merged_df, Bulk_df_Campaign, Bulk_df_AdGroup, Bulk_df_Ad, Bulk_df_Keyword])
    # bulk_df.head()


    # # path to save 
    # path = f'{main_path}\Output'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # bulk_df.to_csv(f'{path}\\New Agent Bulk Upload - {today_date}.csv', index=False)
    # #                     sheet_name='DS Upload Sheet')
    # #                     

    return Sitelink_df, Snippet_df, Call_df, bulk_df