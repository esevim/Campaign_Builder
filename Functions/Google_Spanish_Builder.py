# Import Libraries

import warnings
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
import re, string
from textblob import TextBlob

def main(df, ref_table):
    warnings.filterwarnings('ignore')

    # Getting the todays date
    today_date = date.today().strftime("%m.%d.%y")

    rename_dict = {
        'Agent' : 'Customer', 
        'Act #' : 'Account Number',
        'Landing Page Domain' : 'Website',
        'Geo Target' : 'Geo Target',
        'Market Area' : 'Account', 
        'Lob Exclusion' : 'Lob Exclusion',
        'Banking KW' : 'Banking KWs',
        'Life KW' : 'Life KWs',
        'Spanish KW' : 'Spanish KWs',
        'State' : 'State',
        'Phone Number' : 'Phone Number',
        'Associate ID' : 'Associate ID',
        'City' : 'Location'
    }

    df.rename(mapper = rename_dict, axis=1, inplace = True)

    df['Duration'] = 12

    # Add "Local State" naming to the Account Name
    ref_1 = pd.read_excel(ref_table, sheet_name = 'Account #', index_col='Account')
    df['Customer ID'] = df['Account'].apply(lambda x: ref_1.loc[f'{x}']['Customer ID'])
    df['Account'] = df['Account'].apply(lambda x:'Local State Farm Agents - US - ' + x + ' Region - ' + ref_1.loc[f'{x}']['#'].astype(str))

    # Remove un-wanted text from Customer Name
    df['Customer'] = df['Customer'].str.extract(r"(^.*?(?= - State))")
    # df['Customer']

    # Standardize the Phone Number
    df['Phone Number'] = df['Phone Number'].replace(regex=r'[^0-9.]', value='')
    df['Phone Number'] = '(' + df['Phone Number'].str[:3] + ') ' + df['Phone Number'].str[3:6] + '-' + df['Phone Number'].str[6:]

    # Remove spaces on State Column
    df['State'] = df['State'].str.strip()

    # Create State_long column with full names of states
    states = {"AL":"Alabama", "AK":"Alaska", "AZ":"Arizona", "AR":"Arkansas", "CA":"California", "CO":"Colorado", "CT":"Connecticut", 
            "DC":"Washington DC", "DE":"Delaware", "FL":"Florida", "GA":"Georgia", "HI":"Hawaii", "ID":"Idaho", "IL":"Illinois", 
            "IN":"Indiana", "IA":"Iowa", "KS":"Kansas", "KY":"Kentucky", "LA":"Louisiana", "ME":"Maine", "MD":"Maryland",
            "MA":"Massachusetts", "MI":"Michigan", "MN":"Minnesota", "MS":"Mississippi", "MO":"Missouri", "MT":"Montana",
            "NE":"Nebraska", "NV":"Nevada", "NH":"New Hampshire", "NJ":"New Jersey", "NM":"New Mexico", "NY":"New York", 
            "NC":"North Carolina", "ND":"North Dakota", "OH":"Ohio", "OK":"Oklahoma", "OR":"Oregon", "PA":"Pennsylvania", 
            "RI":"Rhode Island", "SC":"South Carolina", "SD":"South Dakota", "TN":"Tennessee", "TX":"Texas", "UT":"Utah", "VT":"Vermont",
            "VA":"Virginia", "WA":"Washington", "WV":"West Virginia","WI":"Wisconsin", "WY":"Wyoming"}
    
    df["State_long"] = df.State.map(states)


    # Get Check the Website column of '/localagent' is present, if not add.
    def Website_creator(x): 
        website  = str(x['Website'])
        if website[-1:] == '/':
            website = website[:-1] 
        
        if website[-11:] != '/localagent':
            website += '/localagent'
        
        return website

    # Apply URLcreator function to update the Final URL column
    df['Website'] = df.apply(Website_creator, axis=1) 


    Campaign_df = df[['Customer', 
                    'Account Number', 
                    'Duration',
                    'Website',
                    'Geo Target',
                    'Account',
                    'Lob Exclusion',
                    'Banking KWs',
                    'Life KWs',
                    'Spanish KWs',
                    'State',
                    'State_long',
                    'Phone Number',
                    'Location',
                    'Associate ID',
                    'Customer ID'
        ]]


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
            if word == 'Re':
                word = 'RV'
            words.append(word)
            
        res = ', '.join(words)
        return res

    Campaign_df['Lob Exclusion'] = Campaign_df.apply(spell_check, axis=1)


    # ## Sitelink Upload  --------------

    Sitelink_df = Campaign_df.copy()
    ref = pd.read_excel(ref_table, sheet_name='Sitelink') # read Ref table

    def Sitelink_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for['Customer ID'] = row['Customer ID']
            data_for.replace('agentname', row['Customer'], regex=True, inplace=True)
            data_for.replace('00000', row['Account Number'], regex=True, inplace=True)
            data_for.replace('https://www.URL.com/localagent', row['Website'], regex=True, inplace=True)
            data_for['Final URL'] += '/?lang=es'
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Sitelink_df = Sitelink_update(Sitelink_df)


    columns = 'Action	Campaign	Customer ID	Sitelink text	Final URL	Description 1	Description 2'
    columns = columns.split('\t')
    Sitelink_df = pd.DataFrame(data = Sitelink_df, columns=columns)

    # # path to save 
    # path = f'{main_path}\Sitelink Upload'

    # # Checks if path exists
    # if os.path.isdir(main_path) == False:
    #     os.mkdir(main_path)

    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Sitelink_df.to_excel(f'{path}\\New Agent Sitelink Upload - {today_date}.xlsx', 
    #                     sheet_name='Sitelink Upload', 
    #                     index=False)


    # ## Snippet Upload --------------

    Snippet_df = Campaign_df.copy()
    ref = pd.read_excel(ref_table, sheet_name='Snippet') # read Ref table


    def Snippet_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for['Customer ID'] = row['Customer ID']
            data_for.replace('agentname', row['Customer'], regex=True, inplace=True)
            data_for.replace('00000', row['Account Number'], regex=True, inplace=True)
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Snippet_df = Snippet_update(Snippet_df)



    columns = 'Action	Customer ID	Campaign	Extension type	Structured snippet header	Structured Snippet values'
    columns = columns.split('\t')
    Snippet_df = pd.DataFrame(data = Snippet_df, columns=columns)


    # # path to save 
    # path = f'{main_path}\Snippet Upload'

    # # Checks if path exists
    # if os.path.isdir(main_path) == False:
    #     os.mkdir(main_path)

    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Snippet_df.to_excel(f'{path}\\New Agent Snippet Upload - {today_date}.xlsx', 
    #                     sheet_name='Structured Snippet Upload', 
    #                     index=False)

    # ## Call Upload --------------

    Call_df = Campaign_df.copy()
    ref = pd.read_excel(ref_table, sheet_name='Call Template') # read ref table


    def Call_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for['Customer ID'] = row['Customer ID']
            data_for.replace('agentname', row['Customer'], regex=True, inplace=True)
            data_for.replace('00000', row['Account Number'], regex=True, inplace=True)
            data_for['Phone number'] = row['Phone Number']
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Call_df = Call_update(Call_df)


    columns = 'Extension action	Customer ID	Campaign	Phone number	Country code	Use call forwarding	Conversion Action'
    columns = columns.split('\t')
    Call_df = pd.DataFrame(data = Call_df, columns=columns)

    # # path to save 
    # path = f'{main_path}\Call Upload'

    # # Checks if path exists
    # if os.path.isdir(main_path) == False:
    #     os.mkdir(main_path)

    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Call_df.to_excel(f'{path}\\New Agent Call Upload - {today_date}.xlsx', 
    #                     sheet_name='Call Upload', 
    #                     index=False)


    # ## Location Upload  --------------
    # - Agents with multiple Geo Target are excluded
    # - Agents without mileage are excluded

    radius_df = Campaign_df[['Account','Customer','Account Number','Geo Target']].copy()
    ref_1 = pd.read_excel(ref_table, sheet_name='Geo code Coordinates', skiprows=[0,1]) # read Ref table - SF
    ref_2 = pd.read_excel(ref_table, sheet_name = 'Campaign-Template')['Campaign']


    # Create the Campaign Name from reference file.
    # add Agentname and account number
    res = pd.DataFrame(columns=[ref_2.name] + list(radius_df.columns))
    for i, row in radius_df.iterrows():
    #     print(row)
        data_for = ref_2.copy()
        data_for.replace('agentname', row['Customer'], regex=True, inplace=True)
        data_for.replace('00000', row['Account Number'], regex=True, inplace=True)
        data_for = data_for.to_frame()
        data_for['Account Number'] = row['Account Number']
        data_for = data_for.merge(radius_df[i:i+1])
        res = pd.concat([res,data_for], axis=0, ignore_index=True)

    radius_df = res.copy()


    res = pd.DataFrame(columns=['Account', 'Campaign', 'Location', 'Latitude', 'Longitude', 'Location Radius', 'Location Radius Units'])
    for i, x in radius_df.iterrows():
        # Check if the Geo Target column has some content
        if (len(str(x['Geo Target'])) < 1) | (x['Geo Target'] != x['Geo Target']):
            continue
            
        # If there is only Zip code
        if (type(x['Geo Target']) == str):
            Geo_target = x['Geo Target'].split(', ')
        elif (type(x['Geo Target']) == int):
            Geo_target = [str(x['Geo Target'])]
        else:
            Geo_target = list(str(x['Geo Target']))
        
        zip_code_list = []
        miles_list = []
            
        for target in Geo_target:
            target = target.lower()
            # Find zip code for each element of Geo Target.
            # if no zip code found, use the latest found and write it to the list.
            zip_code = re.findall(r'(?<!\d)\d{5}(?!\d)', target)
            if len(zip_code) == 0:
                if len(zip_code_list) == 0:
                    continue
                else:
                    zip_code = [zip_code_list[-1]]
            zip_code_list += zip_code
            
            # Find miles for each element of Geo Target.
            # If no Miles found, write an empty element.
            mile = ['']
            if 'mile radius' in target:
                mile = re.findall(r"(\d+) mile radius", target)
            miles_list += mile    
        
        
        # Create the rows for each element
        for i in range(len(miles_list)):
            zip_code = zip_code_list[i]
            mile = miles_list[i]
            
            res.loc[len(res.index)] = [x['Account'], x['Campaign'], zip_code, '', '', '', '']
            
            lat_long = ref_1[ref_1['Zip'] == int(zip_code)]['Lat_long'].values
            if len(lat_long) > 0:
                lat, long = lat_long[0].split(':')
                if mile != '':
                    res.loc[len(res.index)] = [x['Account'], x['Campaign'], None, lat, long, mile, 'mi']


    radius_df = res.copy()


    radius_df['Row Type'] = radius_df['Location'].apply(lambda x: 'proximity target' if x == None else 'location target')
    radius_df['Action'] = 'create'

    radius_df = radius_df[['Row Type', 'Action', 'Account', 'Campaign', 'Location', 'Latitude', 'Longitude', 'Location Radius','Location Radius Units']]

    # # path to save 
    # path = f'{main_path}\Location Upload'

    # # Checks if path exists
    # if os.path.isdir(main_path) == False:
    #     os.mkdir(main_path)

    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # radius_df.to_excel(f'{path}\\New Agent Radius Target Upload - {today_date}.xlsx', 
    #                     sheet_name='Radius Location Upload', 
    #                     index=False)


    # ## Bulk Upload  --------------
    # ### 1) Campaign Type --------------

    Campaign_Type_df = Campaign_df.copy()

    # Create the Campaign Name from reference file.
    # add Agentname and account number
    res = pd.DataFrame(columns=[ref_2.name] + list(Campaign_Type_df.columns))
    for i, row in Campaign_Type_df.iterrows():
    #     print(row)
        data_for = ref_2.copy()
        data_for.replace('agentname', row['Customer'], regex=True, inplace=True)
        data_for.replace('00000', row['Account Number'], regex=True, inplace=True)
        data_for = data_for.to_frame()
        data_for['Account Number'] = row['Account Number']
        data_for = data_for.merge(Campaign_Type_df [i:i+1])
        res = pd.concat([res,data_for], axis=0, ignore_index=True)
        
    Campaign_Type_df = res.copy()
    # res.shape

    # Full List of Regions
    Excluded_region = 'US-AK | US-AL | US-AR | US-AZ | US-CA | US-CO | US-CT | US-DE | US-FL | US-GA | US-HI | US-IA | US-ID | US-IL | US-IN | US-KS | US-KY | US-LA | US-MA | US-MD | US-ME | US-MI | US-MN | US-MO | US-MS | US-MT | US-NC | US-ND | US-NE | US-NH | US-NJ | US-NM | US-NV | US-NY | US-OH | US-OK | US-OR | US-PA | US-RI | US-SC | US-SD | US-TN | US-TX | US-UT | US-VA | US-VT | US-WA | US-WI | US-WV | US-WY'

    # All Countries except 'US'
    Excluded_country = 'AF|AL|AQ|DZ|AS|AD|AO|AG|AZ|AR|AU|AT|BS|BH|BD|AM|BB|BE|BM|BT|BO|BA|BW|BV|BR|BZ|IO|SB|VG|BN|BG|MM|BI|BY|KH|CM|CA|CV|KY|CF|LK|TD|CL|CN|TW|CX|CC|CO|KM|YT|CG|CD|CK|CR|HR|CY|CZ|BJ|DK|DM|DO|EC|SV|GQ|ET|ER|EE|FO|FK|GS|FJ|FI|FR|GF|PF|TF|DJ|GA|GE|GM|PS|DE|GH|GI|KI|GR|GL|GD|GP|GU|GT|GN|GY|HT|HM|VA|HN|HK|HU|IS|IN|ID|IQ|IE|IL|IT|CI|JM|JP|KZ|JO|KE|KR|KW|KG|LA|LB|LS|LV|LR|LY|LI|LT|LU|MO|MG|MW|MY|MV|ML|MT|MQ|MR|MU|MX|MC|MN|MD|ME|MS|MA|MZ|OM|NA|NR|NP|NL|BQ|CW|AW|SX|BQ|NC|VU|NZ|NI|NE|NG|NU|NF|NO|MP|UM|FM|MH|PW|PK|PA|PG|PY|PE|PH|PN|PL|PT|GW|TL|PR|QA|RE|RO|RU|RW|SH|KN|AI|LC|PM|VC|SM|ST|SA|SN|RS|SC|SL|SG|SK|VN|SI|SO|ZA|ZW|ES|EH|SR|SJ|SZ|SE|CH|TJ|TH|TG|TK|TO|TT|AE|TN|TR|TM|TC|TV|UG|UA|MK|EG|GB|GG|JE|TZ|VI|BF|UY|UZ|VE|WF|WS|YE|ZM|XK|'

    def remove_duplicates_list(x):
        return list(dict.fromkeys(x))


    # Metro Column creator Function
    def Metro_Creator(x):
        geo_target = str(x['Geo Target']).lower() 
        # geo_target = '''15473 10 mile radius, 15236 10 mile radius, 12345, fayette county, ayetste ountay'''

        # Get all ZIP Codes from the text (5 or 5 digit)
        geocodes = re.findall(r'[0-9]{4,5}', geo_target)
        
        # Find all Miles (1 or 3 digit)
        miles = re.findall(r"([0-9]{1,3}) mile", geo_target)

        # Remove each ZIP code from the text
        for code in geocodes:
            geo_target = geo_target.replace(code, '')

        # Remove each Mile radius from text
        for each_mile in miles:
            geo_target = geo_target.replace(each_mile, '')
        
    #     print('a', geo_target)
            
    #     print(geo_target)
        geo_target = geo_target.replace('.', '')
        geo_target = geo_target.replace('miles', '')
        geo_target = geo_target.replace('mile', '')
        geo_target = geo_target.replace('radius', '')
        geo_target = geo_target.replace(' ,', '')
    #     print('b', geo_target)
        
        # remove ', ' if exist and put all in to a list
        
        list_geo_target = [x for x in geo_target.split(', ') if x]
    #     print('c', list_geo_target)
        
    #     print('a')
    #     if x['Geo Target'] == '93312, 93313, 93314, 93311, 93249, 93215, 93250, 93309 20 Mile Radius,':
    #         print('b')
    #         print(list_geo_target)
        
        # Capitalize each Word
        list_geo_target = [string.capwords(item) for item in list_geo_target]
    #     print('d', list_geo_target)
        
        # Add ZIP Codes back
        list_geo_target = list_geo_target + (geocodes)
    #     print('e', list_geo_target)
        
        # print(len(list_geo_target))
        
        
        if (len(list_geo_target) == 1) & (len(miles) == 1) :
            return '' 
            
        list_geo_target = remove_duplicates_list(list_geo_target)
        list_geo_target = ' | '.join(list_geo_target)
    #     print('f', list_geo_target)
    #     print('g', miles)
    #     if len(miles) == 1:
    #         list_geo_target = miles[0] + ' Miles Radius | ' + list_geo_target
        if list_geo_target[:3] == " | ":
            list_geo_target = list_geo_target[3:]
    #         print(list_geo_target)
        list_geo_target = list_geo_target.replace(' |  | ',' | ')
        list_geo_target = list_geo_target.replace(', | ','')
        if list_geo_target == 'Nan':
                list_geo_target = '' 
        return list_geo_target #geocodes, miles, geo_target 


    # Main Function to create all columns for Campaign Type
    def Campaign_row_type(x):
        x['Row Type'] = 'Campaign'
        x['Action'] = 'Create|Pause'
        x['Campaign start date (YYYY-MM-DD)'] = date.today().strftime("%Y-%m-%d")

        year_ahead = (date.today() + relativedelta(months=+x['Duration']))
        x['Campaign end date (YYYY-MM-DD)'] = year_ahead.replace(day=pd.Period(str(year_ahead)).days_in_month).strftime("%Y-%m-%d")

        x['Campaign daily budget'] = 5
        x['Campaign budget type'] = ''
        x['Campaign target network'] = 'Search'
        x['Campaign language target'] = 'ES'
        x['Ad rotation'] = 'Rotate Indefinitely'
        x['Ad delivery method'] = 'Standard'
        x['Label'] = x['State_long']

        to_remove = x['State']
        if to_remove == 'AK':
            x['Excluded region'] = Excluded_region[8:]
        else :
            x['Excluded region'] = Excluded_region.replace(f' | US-{to_remove}', '')

        x['Excluded country'] = Excluded_country
        x['Custom Parameter'] = '{_sfassociateid} = ' + x['Associate ID']
        
    #     if x['Campaign'] == 'DAC State Farm | Standard | Brand | Exact | Abir M Pulskamp | 226-01-7700-05320':
    #         print(x['Geo Target'])
            
    #     x['Metro'] = Metro_Creator(x)
        x['Metro'] = ''
        return x[['Row Type', 'Action', 'Account', 'Campaign', 
                'Campaign start date (YYYY-MM-DD)', 'Campaign end date (YYYY-MM-DD)', 
                'Campaign daily budget', 'Campaign budget type', 'Campaign target network',
                'Campaign language target', 'Metro', 'Ad rotation', 'Ad delivery method',
                'Label', 'Excluded region', 'Excluded country', 'Custom Parameter']]

    Bulk_df_Campaign = Campaign_Type_df.apply(Campaign_row_type, axis=1)
    # Bulk_df_Campaign.head(1)


    # ### 2) Ad Group Type  --------------

    ref = pd.read_excel(ref_table, sheet_name='AdGroup-Template')[['Campaign', 'Ad Group']].dropna(how='all') # read ref table
    ref2 = pd.read_excel(ref_table, sheet_name='AdGroup-Template')[['INACTIVE KWs For anyone']].dropna(how='all') # read Ref table - SF
    ref3 = pd.read_excel(ref_table, sheet_name='AdGroup-Template')[['KWs','States','Until Date']].dropna(how='all') # read Ref table - SF
    ref3['Until Date'] = pd.to_datetime(ref3['Until Date'])
    AdGroup_df = Campaign_df.copy()

    # Check the main data and add all LOB Exclusions to a main list.
    def isNaN(string):
        return string != string


    def LOB(x):
        if (len(str(x['Lob Exclusion'])) < 1) | isNaN(x['Lob Exclusion']):
            Lob_Exclusions = None
            return Lob_Exclusions
        
        Lob_Exclusions = x['Lob Exclusion'].split(', ')
        
        # Special cases for LOB's
        Lob_Exclusions = ['EV' if KW == 'Electric Vehicle' else KW for KW in Lob_Exclusions]
        Lob_Exclusions = ['Farm_Ranch' if KW == 'Farm/Ranch' else KW for KW in Lob_Exclusions]
        Lob_Exclusions = ['RV' if KW == 'Re' else KW for KW in Lob_Exclusions]
        
        if 'Financial' in Lob_Exclusions:
            Lob_Exclusions.extend(['Financial Services'])
        if 'Financial Services' in Lob_Exclusions:
            Lob_Exclusions.extend(['Financial'])
        if 'Auto Quote' in Lob_Exclusions:
            Lob_Exclusions.extend(['Auto'])
        if 'Auto' in Lob_Exclusions:
            Lob_Exclusions.extend(['Auto Quote'])
        
        if 'Mobile Home' in Lob_Exclusions:
            Lob_Exclusions.extend(['Mobile Home', 'Trailer'])
        if 'Mobile' in Lob_Exclusions:
            Lob_Exclusions.extend(['Mobile', 'Trailer', 'Motorhome', 'RV'])
        if 'Commercial Auto' in Lob_Exclusions:
            Lob_Exclusions.extend(['Business'])
        if 'Business' in Lob_Exclusions:
            Lob_Exclusions.extend(['Commercial Auto'])
            
        if x['Banking KWs'] == 'No':
            Lob_Exclusions.append('Banking')
        if x['Life KWs'] == 'No':
            Lob_Exclusions.append('Life')
        if x['Spanish KWs'] == 'No':
            Lob_Exclusions.append('Spanish') 
        
        # Special Cases for Each state and date
        Lob_Exclusions += list(ref2.iloc[:,0].values)
        
        for i in range(ref3.shape[0]):
            Ref_KW = ref3.iloc[i]['KWs']
            Ref_State = ref3.iloc[i]['States']
            
            # If Until Date is empty, take it as tomorrows day, so the below function will work
            if pd.isnull(ref3.iloc[i]['Until Date']):
                Ref_date = date.today() + relativedelta(days=+1)
            else: 
                Ref_date = ref3.iloc[i]['Until Date']
            
            if (Ref_KW not in x['Lob Exclusion'].split(', ')) & (x['State'] in Ref_State) & (Ref_date > date.today()):
    #             print(Lob_Exclusions, '\n', Ref_KW, '\n', Ref_State, '\n', x['Lob Exclusion'])
                Lob_Exclusions.append(Ref_KW)
            
        # Add 'Agent' to the end of each LOB
        _ = Lob_Exclusions.copy()
        for Lob in Lob_Exclusions:
            _.extend([f'{Lob} Agent'])
        Lob_Exclusions = _

        # Remove dublicates
        Lob_Exclusions = list(dict.fromkeys(Lob_Exclusions))
        
        return Lob_Exclusions
        
    AdGroup_df['Lob Exclusion'] = AdGroup_df.apply(LOB, axis=1)


    def AdGroup_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for.replace('agentname', row['Customer'], regex=True, inplace=True)
            data_for.replace('00000', row['Account Number'], regex=True, inplace=True)
            
            data_for['Action'] = 'Create'
            i = 0
            for row2 in data_for['Ad Group']:
                if row2.split(' - ')[2] in row['Lob Exclusion']:
    #                 print(row2.split(' - ')[1], row['Lob Exclusion'])
                    data_for['Action'][i] = 'Create|Pause'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        res['Row Type'] = 'Ad Group'
        res['Ad group search max CPC'] = 3
        
        return res

    Bulk_df_AdGroup = AdGroup_update(AdGroup_df)


    # ### 3) Ad Type --------------

    ref = pd.read_excel(ref_table, sheet_name='Ad-Template') # read ref table
    Ad_df = Campaign_df.copy()


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
                if row2.split(' - ')[2] in row['Lob Exclusion']:
    #                 print(row2.split(' - ')[1], row['Lob Exclusion'])
    #                 data_for['Action'][i] = 'Create|Pause'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            
            # Ad Description and Ad Title Columns Lenght Edit
            for column in ['Ad description line 1','Ad description line 2', 'Ad description line 3', 'Ad description line 4']:
                i = 0
                for row3 in data_for[column]:
                    if row3 != row3:
                        i+=1
                        continue
                    elif len(row3) + len(row['Customer']) - agentname_len > 90:
                        data_for[column][i] = data_for[column][i].replace('agentname', row['Customer First Name'])
    #                     print(len(data_for[column][i]))
    #                     if len(data_for[column][i]) > 90:
    #                         print(row['Customer'])
    #                         print(data_for[column][i])
                    else: 
                        data_for[column][i] = data_for[column][i].replace('agentname', row['Customer'])
                    i += 1

            for column in ['Ad title', 'Ad title 2', 'Ad title 3','Ad title 4', 'Ad title 5', 'Ad title 6', 'Ad title 7', 
                        'Ad title 8', 'Ad title 9', 'Ad title 10', 'Ad title 11', 'Ad title 12','Ad title 13', 'Ad title 14', 'Ad title 15']:
                i = 0
                for row3 in data_for[column]:
    #                 print('Agent Name:', row['Customer'])
                    
                    if row3 != row3:
                        i+=1
                        continue
                    if len(row3) + len(row['Customer']) - agentname_len > 30:
                        data_for[column][i] = data_for[column][i].replace('agentname', row['Customer First Name'])
                        
    #                     if len(data_for[column][i]) > 30:
    #                         print(row['Customer'])
    #                         print(data_for[column][i])
                    else: 
                        data_for[column][i] = data_for[column][i].replace('agentname', row['Customer'])
                    i += 1
        

            data_for.replace('agentname', row['Customer'], regex=True, inplace=True)
            data_for.replace('00000', row['Account Number'], regex=True, inplace=True)
            data_for['Ad landing page'].replace('https://www.URL.com/localagent', row['Website'], regex=True, inplace=True)
            data_for['Ad landing page'] += '/?lang=es'
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
            
        res['Row Type'] = 'Ad'
        return res


    agentname_len = len('agentname')

    Bulk_df_Ad = Ad_update(Ad_df)

    # ### 4) Keyword Type  --------------

    ref = pd.read_excel(ref_table, sheet_name='Keyword-Template') # read ref table
    Keyword_df = Campaign_df.copy()


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
                
                if row2.split(' - ')[2] in row['Lob Exclusion']:
    #                 data_for['Action'][i] = 'Create|Pause'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            
            data_for.replace('agentname', row['Customer'], regex=True, inplace=True)
            data_for.replace('00000', row['Account Number'], regex=True, inplace=True)
            
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        res['Row Type'] = 'Keyword'

        res.rename(columns={'Max CPC':'Keyword max CPC'}, inplace=True)
        
        return res


    Bulk_df_Keyword = Keyword_update(Keyword_df).rename(columns={'Max CPC':'Keyword max CPC',
                                            'Criterion Type':'Keyword match type'})
    

    # ### 5) Merge four subtable to one --------------

    columns = list(Bulk_df_Campaign.columns) + list(Bulk_df_AdGroup.columns) + list(Bulk_df_Ad.columns) + list(Bulk_df_Keyword.columns)
    columns = list(dict.fromkeys(columns))

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
        'Ad type',
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
        'Keyword max CPC',
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


    bulk_df = pd.concat([Merged_df, Bulk_df_Campaign, Bulk_df_AdGroup, Bulk_df_Ad, Bulk_df_Keyword]).reset_index(drop=True)

    res = pd.DataFrame(columns=columns)
    for campaign in bulk_df['Campaign'].unique():
        data_for = bulk_df[bulk_df['Campaign'] == campaign]
        data_for['Account'] = data_for['Account'].fillna(method='ffill')
        res = pd.concat([res,data_for], axis=0, ignore_index=True)
        
    bulk_df = res.copy()

    # # path to save 
    # path = f'{main_path}'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # bulk_df.to_csv(f'{path}\\New Agent Bulk Upload - {today_date}.csv', 
    # #                     sheet_name='DS Upload Sheet', 
    #                     index=False)

    return Sitelink_df, Snippet_df, Call_df, radius_df, bulk_df