# # Notes
# - 3.28.2022 - Tool Created
# - 4.08.2022 - Bulk upload Created
# - 6.01.2022 - Small errors fixed, moved to the VM automation
# - 9.13.2022 - New Editor Location Upload Format created. MEtro column on Bulk Upload been emptied.<br>
# 'Customer ID' column added to small upload files as per region.
# - 10.20.2022 - Special Ad cases for LOB and Regions is added
# - 11.11.2022 - Streamlit Version created

# Import Libraries
import os
import re
import string
from datetime import date

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta


def main(df, Ref_df):
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
    source_df = df[['Customer', 
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
                    'Phone Number',
                    'Associate ID',
                    'Location',       
        ]]

    # Add "Local State" naming to the Account Name
    ref = pd.read_excel(Ref_df, sheet_name='Account #', index_col='Account')
    source_df['Customer ID'] = source_df['Account'].apply(lambda x: ref.loc[f'{x}']['Customer ID'])
    source_df['Account'] = source_df['Account'].apply(lambda x:'Local State Farm Agents - US - ' + x + ' Region - ' + ref.loc[f'{x}']['#'].astype(str))

    # Remove un-wanted text from Customer Name
    source_df['Customer'] = source_df['Customer'].str.extract(r"(^.*?(?= - State))")

    # Standardize the Phone Number
    source_df['Phone Number'] = source_df['Phone Number'].replace(regex=r'[^0-9.]', value='')
    source_df['Phone Number'] = '(' + source_df['Phone Number'].str[:3] + ') ' + source_df['Phone Number'].str[3:6] + '-' + source_df['Phone Number'].str[6:]

    # Remove spaces on State Column
    source_df['State'] = source_df['State'].str.strip()

    states = {"AL":"Alabama", "AK":"Alaska", "AZ":"Arizona", "AR":"Arkansas", "CA":"California", "CO":"Colorado", "CT":"Connecticut", 
            "DC":"Washington DC", "DE":"Delaware", "FL":"Florida", "GA":"Georgia", "HI":"Hawaii", "ID":"Idaho", "IL":"Illinois", 
            "IN":"Indiana", "IA":"Iowa", "KS":"Kansas", "KY":"Kentucky", "LA":"Louisiana", "ME":"Maine", "MD":"Maryland",
            "MA":"Massachusetts", "MI":"Michigan", "MN":"Minnesota", "MS":"Mississippi", "MO":"Missouri", "MT":"Montana",
            "NE":"Nebraska", "NV":"Nevada", "NH":"New Hampshire", "NJ":"New Jersey", "NM":"New Mexico", "NY":"New York", 
            "NC":"North Carolina", "ND":"North Dakota", "OH":"Ohio", "OK":"Oklahoma", "OR":"Oregon", "PA":"Pennsylvania", 
            "RI":"Rhode Island", "SC":"South Carolina", "SD":"South Dakota", "TN":"Tennessee", "TX":"Texas", "UT":"Utah", "VT":"Vermont",
            "VA":"Virginia", "WA":"Washington", "WV":"West Virginia","WI":"Wisconsin", "WY":"Wyoming"}
    
    source_df["State_long"] = source_df.State.map(states)

    # Dictionary from full State name to Abbreviations
    State_abbreviations = {v: k for k, v in states.items()}

    List_of_States = list(State_abbreviations.keys())

    # Get Check the Website column of '/localagent' is present, if not add.
    def Website_creator(x): 
        website  = str(x['Website'])
        if website[-1:] == '/':
            website = website[:-1] 
        
        if website[-11:] != '/localagent':
            website += '/localagent'
        
        return website

    source_df['Website'] = source_df.apply(Website_creator, axis=1) # Apply URLcreator function to update the Final URL column

    # Get the Campaign Names of specified Account Numbers and Names
    def Campaign_name_creator(account_numbers, customer_names):
        brandnb = ['Brand', 'NB']
        bmmexact = ['Exact', 'BMM']
        res = pd.DataFrame(columns=['Campaign', 'Brand Type', 'Match Type'])
        loop_df = pd.concat([account_numbers, customer_names], axis=1)

        for agent_id, name in loop_df.values:   
            for brand in brandnb:
                for bmm in bmmexact:
                    row = pd.Series({'Campaign':f'DAC State Farm | Standard | {brand} | {bmm} | {name} | {agent_id}',
                                    'Brand Type':brand, 
                                    'Match Type':bmm,
                                    'Account Number': agent_id})
                    res = res.append(row, ignore_index=True)
        return res

    Campaign_df = Campaign_name_creator(source_df['Account Number'], source_df['Customer'])
    Campaign_df = Campaign_df.merge(source_df, on='Account Number')

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

    Campaign_df['Lob Exclusion'] = Campaign_df.apply(spell_check, axis=1)

    ## Snippet Upload --------------
    Snippet_df = Campaign_df[['Customer ID','Campaign']].copy()
    Snippet_df['Action'] = 'add'
    Snippet_df['Extension type'] = 'Structured snippet extension'
    Snippet_df['Structured snippet header'] = 'Insurance coverage'
    Snippet_df['Structured Snippet values'] = 'Auto;Home;Renters;Condo'

    Snippet_df = Snippet_df[['Action','Customer ID','Campaign','Extension type','Structured snippet header','Structured Snippet values']]

    # # path to save 
    # path = f'{main_path}\Snippet Upload'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Snippet_df.to_excel(f'{path}\\New Agent Snippet Upload - {today_date}.xlsx', 
    #                     sheet_name='Structured Snippet Upload', 
    #                     index=False)

    ## Sitelink Upload --------------
    Sitelink_df = Campaign_df.copy()
    ref = pd.read_excel(Ref_df, sheet_name='Sitelink') # read Ref table - SF
    Sitelink_df = ref.merge(Sitelink_df, on=['Brand Type', 'Match Type'])   # get Descriptions and other text from Ref table - SF.
    Sitelink_df.sort_values(['Campaign', 'Brand Type','Match Type'], inplace=True) #Sort

    # Get convention from Sitelink Text column and add to website URL
    def URLcreator(x): 
        sitelink = str(x['Sitelink text'])
        website  = str(x['Website'])
        
        sitelink = sitelink.replace(' ', '-').lower()
        sitelink = sitelink.replace('owners', '')
        sitelink = sitelink.replace('-quote', '')

        website = website +'/'+ sitelink
        return website

    Sitelink_df['Final URL'] = Sitelink_df.apply(URLcreator, axis=1) # Apply URLcreator function to update the Final URL column
    Sitelink_df.sort_values(['Account Number', 'Campaign'], inplace=True, ignore_index=True) # Sort
    Sitelink_df = Sitelink_df[['Action','Campaign','Customer ID', 'Sitelink text','Final URL','Description 1', 'Description 2']]

    # # path to save 
    # path = f'{main_path}\Sitelink Upload'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Sitelink_df.to_excel(f'{path}\\New Agent Sitelink Upload - {today_date}.xlsx', 
    #                     sheet_name='Sitelink Upload', 
    #                     index=False)

    ## Location Upload -------------------------
    # - Agents with multiple Geo Target are excluded
    # - Agents without mileage are excluded

    Radius_df = Campaign_df[['Account','Campaign','Geo Target']].copy()
    ref = pd.read_excel(Ref_df, sheet_name='Geo code Coordinates', skiprows=[0,1]) # read Ref table - SF

    res = pd.DataFrame(columns=['Account', 'Campaign', 'Location', 'Latitude', 'Longitude', 'Location Radius', 'Location Radius Units'])
    for i, x in Radius_df.iterrows():
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
    #     print('5',zip_code_list)
    #     print('6',miles_list)    
        
        
        # Create the rows for each element
        for i in range(len(miles_list)):
            zip_code = zip_code_list[i]
            mile = miles_list[i]
            
            res.loc[len(res.index)] = [x['Account'], x['Campaign'], zip_code, '', '', '', '']
            
            lat_long = ref[ref['Zip'] == int(zip_code)]['Lat_long'].values
            if len(lat_long) > 0:
                lat, long = lat_long[0].split(':')
    #             print('4', lat, long)
    #             print([x['Account'], x['Campaign'], None, lat, long, mile, 'mi'])
                if mile != '':
                    res.loc[len(res.index)] = [x['Account'], x['Campaign'], None, lat, long, mile, 'mi']
    #             print(res)
    #         print([x['Account'], x['Campaign'], zip_code[0], '', '', '', ''])      

    Radius_df = res.copy()
    Radius_df['Row Type'] = Radius_df['Location'].apply(lambda x: 'proximity target' if x == None else 'location target')
    Radius_df['Action'] = 'create'

    Radius_df = Radius_df[['Row Type', 'Action', 'Account', 'Campaign', 'Location', 'Latitude', 'Longitude', 'Location Radius','Location Radius Units']]

    # # path to save 
    # path = f'{main_path}\Location Upload'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # radius_df.to_excel(f'{path}\\New Agent Radius Target Upload - {today_date}.xlsx', 
    #                     sheet_name='Radius Location Upload', 
    #                     index=False)

    ### Call Upload ---------------
    Call_df = Campaign_df[['Customer ID', 'Campaign','Phone Number']].copy()
    Call_df['Extension action'] = 'Create new'
    Call_df['Country code'] = 'US'
    Call_df['Use call forwarding'] = 'yes'
    Call_df['Conversion Action'] = 'Call from ads (30 seconds) MCC'

    Call_df = Call_df[['Extension action','Customer ID', 'Campaign', 'Phone Number','Country code', 'Use call forwarding', 'Conversion Action' ]]

    # # path to save 
    # path = f'{main_path}\Call Upload'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # call_df.to_excel(f'{path}\\New Agent Call Upload - {today_date}.xlsx', 
    #                     sheet_name='Call Upload', 
    #                     index=False)

    ### BULK UPLOAD -------------------
    ## 1) Campaign Type ---------------

    Campaign_Type_df = Campaign_df.copy()

    # Full List of Regions
    Excluded_region = 'US-AK | US-AL | US-AR | US-AZ | US-CA | US-CO | US-CT | US-DE | US-FL | US-GA | US-HI | US-IA | US-ID | US-IL | US-IN | US-KS | US-KY | US-LA | US-MA | US-MD | US-ME | US-MI | US-MN | US-MO | US-MS | US-MT | US-NC | US-ND | US-NE | US-NH | US-NJ | US-NM | US-NV | US-NY | US-OH | US-OK | US-OR | US-PA | US-RI | US-SC | US-SD | US-TN | US-TX | US-UT | US-VA | US-VT | US-WA | US-WI | US-WV | US-WY'

    # All Countries except 'US'
    Excluded_country = 'AF|AL|AQ|DZ|AS|AD|AO|AG|AZ|AR|AU|AT|BS|BH|BD|AM|BB|BE|BM|BT|BO|BA|BW|BV|BR|BZ|IO|SB|VG|BN|BG|MM|BI|BY|KH|CM|CA|CV|KY|CF|LK|TD|CL|CN|TW|CX|CC|CO|KM|YT|CG|CD|CK|CR|HR|CY|CZ|BJ|DK|DM|DO|EC|SV|GQ|ET|ER|EE|FO|FK|GS|FJ|FI|FR|GF|PF|TF|DJ|GA|GE|GM|PS|DE|GH|GI|KI|GR|GL|GD|GP|GU|GT|GN|GY|HT|HM|VA|HN|HK|HU|IS|IN|ID|IQ|IE|IL|IT|CI|JM|JP|KZ|JO|KE|KR|KW|KG|LA|LB|LS|LV|LR|LY|LI|LT|LU|MO|MG|MW|MY|MV|ML|MT|MQ|MR|MU|MX|MC|MN|MD|ME|MS|MA|MZ|OM|NA|NR|NP|NL|BQ|CW|AW|SX|BQ|NC|VU|NZ|NI|NE|NG|NU|NF|NO|MP|UM|FM|MH|PW|PK|PA|PG|PY|PE|PH|PN|PL|PT|GW|TL|PR|QA|RE|RO|RU|RW|SH|KN|AI|LC|PM|VC|SM|ST|SA|SN|RS|SC|SL|SG|SK|VN|SI|SO|ZA|ZW|ES|EH|SR|SJ|SZ|SE|CH|TJ|TH|TG|TK|TO|TT|AE|TN|TR|TM|TC|TV|UG|UA|MK|EG|GB|GG|JE|TZ|VI|BF|UY|UZ|VE|WF|WS|YE|ZM|XK|'

    def remove_duplicates_list(x):
        return list(dict.fromkeys(x))

    # - Regex detection is good
    # - need to decide if want to show both address and zip code on the same line
    # - line 960 has some weird problem.
    
    # Main Function to create all columns for Campaign Type
    def Campaign_row_type(x):
        x['Row Type'] = 'Campaign'
        x['Action'] = 'Create|Pause'
        x['Campaign start date (YYYY-MM-DD)'] = date.today().strftime("%Y-%m-%d")

        year_ahead = (date.today() + relativedelta(months=+x['Duration']))
        x['Campaign end date (YYYY-MM-DD)'] = year_ahead.replace(day=pd.Period(str(year_ahead)).days_in_month).strftime("%Y-%m-%d")

        x['Campaign daily budget'] = 5
        x['Campaign target network'] = 'Search'
        x['Campaign language target'] = 'EN'
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
                'Campaign daily budget', 'Campaign target network',
                'Campaign language target', 'Metro', 'Ad rotation', 'Ad delivery method',
                'Label', 'Excluded region', 'Excluded country', 'Custom Parameter']]
                
    Bulk_df_Campaign = Campaign_Type_df.apply(Campaign_row_type, axis=1)

    ## 2) Ad Group Type ---------------

    Ad_Group_Type_df = Campaign_df.copy()
    ref = pd.read_excel(Ref_df, sheet_name='Bulk-Ad Group')[['Brand Type', 'Match Type', 'LOB']] # read Ref table - SF
    ref2 = pd.read_excel(Ref_df, sheet_name='Bulk-Ad Group')[['INACTIVE KWs For anyone']].dropna(how='all') # read Ref table - SF
    ref3 = pd.read_excel(Ref_df, sheet_name='Bulk-Ad Group')[['KWs','States','Until Date']].dropna(how='all') # read Ref table - SF
    ref3['Until Date'] = pd.to_datetime(ref3['Until Date'])

    # Add Ad Group to the Campaign Data
    Merged_df = Ad_Group_Type_df.merge(ref, on=['Brand Type', 'Match Type'])

    # Get each item on each row of ref3 table
        # if Until date is empty, mark it as tomorrow, so it will run on the following if statement
    # if ref_KW not int current Lob Exclusions,

    def isNaN(string):
        return string != string

    def Ref3_LOB_Exclusions(x):
        
        # Check if LOB Exclusions is empty
        if (len(str(x['Lob Exclusion'])) < 1) | isNaN(x['Lob Exclusion']):
            Lob_Exclusions = None
            return Lob_Exclusions
        
    #     print(x['Lob Exclusion'])
        Lob_Exclusions = x['Lob Exclusion'].split(', ')

        # Special Exclusion cases for LOB's
        if 'Commercial Auto' in Lob_Exclusions:
            Lob_Exclusions.extend(['Business'])
        if 'Mobile Home' in Lob_Exclusions:
            Lob_Exclusions.extend(['Mobile', 'Trailer', 'Motorhome', 'RV'])
        if 'Mobile' in Lob_Exclusions:
            Lob_Exclusions.extend(['RV', 'Trailer', 'Motorhome'])
            
    #     if 'Liability' in Lob_Exclusions:
    #         Lob_Exclusions.extend(['Liability Agent'])
        if 'Business' in Lob_Exclusions:
            Lob_Exclusions.extend(['Commercial Auto'])
        
    #     Lob_Exclusions = x['Lob Exclusion']

        # Special Cases for Each state and date
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
    #             print(type(Ref_KW))
    #             print(Lob_Exclusions, '\n', x['Campaign'])
    #     print(Lob_Exclusions)
        return Lob_Exclusions

    Merged_df['Lob Exclusion'] = Merged_df.apply(Ref3_LOB_Exclusions, axis=1)

    # Check the main data and add all LOB Exclusions to a main list.
    # Check ref2 table and add excluded LOB's to the list
    def LOB(x):
        
    #     print(len(x['Lob Exclusion']))
    #     print(x['Lob Exclusion'])
        if (len(str(x['Lob Exclusion'])) < 1) or x['Lob Exclusion'] == None:
            Lob_Exclusions = None
            return Lob_Exclusions
        
    #     print(x['Lob Exclusion'])
        Lob_Exclusions = x['Lob Exclusion']
        Lob_Exclusions = Lob_Exclusions + list(ref2.iloc[:,0].values)
    
        if x['Banking KWs'] == 'No':
            Lob_Exclusions.append('Banking')
        if x['Life KWs'] == 'No':
            Lob_Exclusions.append('Life')
        if x['Spanish KWs'] == 'No':
            Lob_Exclusions.append('Spanish') 

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
        
        # Add 'Agent' to the end of each LOB
        _ = Lob_Exclusions.copy()
        for Lob in Lob_Exclusions:
            _.extend([f'{Lob} Agent'])
        Lob_Exclusions = _
            
            
        Lob_Exclusions = list(dict.fromkeys(Lob_Exclusions))
    #     print(Lob_Exclusions)
        
        return Lob_Exclusions
        
    Merged_df['Lob Exclusion'] = Merged_df.apply(LOB, axis=1)

    # Combine Brand Type, Ad Group Name and Match Type
    def Ad_Group_Type(x):
        
        if (len(str(x['Lob Exclusion'])) < 1) or x['Lob Exclusion'] == None:
            return None
        
        if x['Match Type'] == 'Exact':
            match_type = 'E'
        elif x['Match Type'] == 'BMM':
            match_type = 'B'
        
        res = x['Brand Type'] + ' - ' + x['LOB'] + ' - ' + match_type
        if x['LOB'] in x['Lob Exclusion']:
            res = 'INACTIVE LOB - ' + res    
        return res

    Merged_df['Ad group'] = Merged_df.apply(Ad_Group_Type, axis=1)

    # Create Action Column as per the Ad Group
    def Action(x):
        if x['Lob Exclusion'] == None:
            res = 'Create'
            return res
        
        if x['Ad group'][:8] == 'INACTIVE':
            res = 'Create|Pause'
        else:
            res = 'Create'
        return res

    Merged_df['Action'] = Merged_df.apply(Action, axis=1)

    Bulk_df_ad_group = Merged_df[['Action','Account','Campaign','Ad group', 'LOB']]
    Bulk_df_ad_group['Row Type'] = 'Ad Group'
    Bulk_df_ad_group['Ad group search max CPC'] = 3

    ## 3) Ad Type ---------------

    Ad_Type_df = Campaign_df[['Campaign', 'Brand Type', 'Match Type', 'Account Number', 'Customer', 'Website','State_long']].copy()
    Ad_Type_df = Ad_Type_df.merge(Bulk_df_ad_group[['Account', 'Campaign', 'Ad group', 'LOB']], on=['Campaign'])

    Ad_Type_df['Customer First Name'] = Ad_Type_df['Customer'].str.split(' ',expand=True)[0]

    ref = pd.read_excel(Ref_df, sheet_name='Bulk-Ad') # read Ref table - SF

    ref_1 = pd.read_excel(Ref_df, sheet_name='Bulk-Ad Overwrite Rules') # read Ref table - SF
    ref_1["State_long"] = ref_1['States to make this changes'].map(states)
    ref_1.drop(columns='States to make this changes', inplace=True)

    filter_rule = (Ad_Type_df['State_long'].isin(ref_1['State_long'])) & (Ad_Type_df['LOB'].isin(ref_1['LOB']))
    Merged_df1 = Ad_Type_df[filter_rule].merge(ref_1, on=['Brand Type', 'Match Type', 'LOB','State_long'])
    Merged_df2 = Ad_Type_df[~filter_rule].merge(ref, on=['Brand Type', 'Match Type', 'LOB'])
    Merged_df = pd.concat([Merged_df1, Merged_df2], ignore_index=True)

    def Agentname(x):
        
        for column in columns:
            if 'agentname' in str(x[column]):
                if column in ['Ad description line 1', 'Ad description line 2', 'Ad description line 3', 'Ad description line 4']:
                    if len(x[column]) + len(x['Customer']) - agentname_len > 90:
                        # Using agents' first name instead of full name to be able to stay under 90 characters
                        x[column] = x[column].replace('agentname', x['Customer First Name']) 
                    else:
                        x[column] = x[column].replace('agentname', x['Customer'])
                    
                elif column in ['Ad title', 'Ad title 2', 'Ad title 3', 'Ad title 4', 'Ad title 5', 'Ad title 6', 'Ad title 7', 'Ad title 8',
                'Ad title 9', 'Ad title 10', 'Ad title 11', 'Ad title 12', 'Ad title 13', 'Ad title 14', 'Ad title 15']:
                    if len(x[column]) + len(x['Customer']) - agentname_len > 30:
                        # Using agents' first name instead of full name to be able to stay under 30 characters
                        x[column] = x[column].replace('agentname', x['Customer First Name']) 
                    else:
                        x[column] = x[column].replace('agentname', x['Customer'])
                                            
                else:
                    x[column] = x[column].replace('agentname', x['Customer'])
        return x

    agentname_len = len('agentname')

    columns = Merged_df.columns

    Merged_df = Merged_df.apply(Agentname, axis=1)

    def Final_URL(x):
        return x.replace('https://www.URL.com/localagent', x['Website'], regex=True)

    Merged_df = Merged_df.apply(Final_URL, axis=1)

    def location(x):
        for column in columns:
            if '[location]' in str(x[column]):
                if column in ['Ad description line 1', 'Ad description line 2', 'Ad description line 3', 'Ad description line 4']:
                    # Using State Abbreviations instead of State Names if characters are longer than 90
                    if len(x[column]) + len(x['State_long']) - location_len > 90:
    #                     print(x[column], len(x[column]))
                        x[column] = x[column].replace('[location]', State_abbreviations[x['State_long']])              
    #                     print(x[column])
                    else :
    #                     print(x[column], len(x[column]))
                        x[column] = x[column].replace('[location]', x['State_long'])
    #                     print(x[column])
            
                elif column in ['Ad title', 'Ad title 2', 'Ad title 3', 'Ad title 4', 'Ad title 5', 'Ad title 6', 'Ad title 7', 
                'Ad title 8', 'Ad title 9', 'Ad title 10', 'Ad title 11', 'Ad title 12', 'Ad title 13', 'Ad title 14', 'Ad title 15']:
                    # Using State Abbreviations instead of State Names if characters are longer than 30
                    if len(x[column]) + len(x['State_long']) - location_len > 30:
    #                     print(x[column])
                        x[column] = x[column].replace('[location]', State_abbreviations[x['State_long']])
    #                     print(x[column])
                    else :
    #                     print(x[column])
                        x[column] = x[column].replace('[location]', x['State_long'])
    #                     print(x[column])
        return x

    location_len = len('[location]')

    columns = Merged_df.columns

    Merged_df = Merged_df.apply(location, axis=1)

    Merged_df['Row Type'] = 'Ad'
    Merged_df['Action'] = 'Create'

    Bulk_df_ad = Merged_df[['Row Type','Action','Account','Campaign','Ad group', 'Ad type','Ad title', 'Ad title 2', 'Ad title 3', 'Ad title 4', 'Ad title 5', 'Ad title 6', 'Ad title 7',
                        'Ad title 8', 'Ad title 9', 'Ad title 10', 'Ad title 11', 'Ad title 12','Ad title 13', 'Ad title 14', 'Ad title 15', 'Ad description line 1','Ad description line 2', 
                        'Ad description line 3', 'Ad description line 4','Ad path field 1', 'Ad path field 2', 'Ad landing page','Label']]

    ## 4) Keyword Type ---------------

    Keyword_Type_df = Campaign_df[['Campaign', 'Brand Type', 'Match Type']].copy()
    Keyword_Type_df = Keyword_Type_df.merge(Bulk_df_ad_group[['Account', 'Campaign', 'Ad group', 'LOB']], on=['Campaign'])
    ref = pd.read_excel(Ref_df, sheet_name='Bulk-Keyword') # read Ref table - SF

    # Add Keywords to the Campaign Data
    Merged_df = Keyword_Type_df.merge(ref, on=['Brand Type', 'Match Type', 'LOB'])

    Bulk_df_Keyword = Merged_df.copy()
    Bulk_df_Keyword['Row Type'] = 'Keyword'
    Bulk_df_Keyword['Action'] = 'Create'

    ## 5) Merge four Subtable to One -----------------

    columns = 'Row Type	Action	Advertiser	Account	Campaign	Advertiser bid strategy	Campaign start date (YYYY-MM-DD)	Campaign end date (YYYY-MM-DD)	Campaign daily budget	Campaign monthly budget	Campaign lifetime budget	Campaign budget type	Campaign target network	Desktop bid adjustment %	Mobile bid adjustment %	Tablet bid adjustment %	Campaign language target	Country	Region	Metro	City	Ad rotation	Ad delivery method	Ad group	Ad group language	Ad group search max CPC	Ad group content max CPC	Ad group native max CPC	Ad group start date (YYYY-MM-DD)	Ad group end date (YYYY-MM-DD)	Ad ID	Ad name	Ad type	Ad title	Ad title 2	Ad title 3	Ad title 4	Ad title 5	Ad title 6	Ad title 7	Ad title 8	Ad title 9	Ad title 10	Ad title 11	Ad title 12	Ad title 13	Ad title 14	Ad title 15	Ad description line 1	Ad description line 2	Ad description line 3	Ad description line 4	Ad path field 1	Ad path field 2	Ad landing page	Ad display URL	Mobile display URL	Ad mobile landing page	Ad device preference	Ad Sponsored By	Keyword ID	Keyword	Keyword match type	Keyword max CPC	Keyword landing page	Keyword min strat bid	Keyword max strat bid	Label	Label ID	Keyword Param 2	Keyword Param 3	URL Template	Excluded region	Excluded country	Custom Parameter'
    columns = columns.split('\t')

    Merged_df = pd.DataFrame(columns=columns)

    Bulk_df = pd.concat([Merged_df, Bulk_df_Campaign, Bulk_df_ad_group, Bulk_df_ad, Bulk_df_Keyword]).iloc[:,:-3]

    return Snippet_df, Sitelink_df, Radius_df, Call_df, Bulk_df


