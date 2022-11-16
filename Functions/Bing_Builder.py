# # Notes
# - 11.16.2022 - Streamlit Version created

# Import Libraries
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
import re
from textblob import TextBlob

def main(df, ref_table):
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
    ref_1 = pd.read_excel(ref_table, sheet_name = 'Account #')
    df['Account'] = 'Local State Farm Agents - ' + str(ref_1['#'][0])
    df['Customer ID'] = str(ref_1['Customer ID'][0])



    # Remove un-wanted text from Customer Name
    df['Customer'] = df['Customer'].str.extract(r"(^.*?(?= - State))")
    # df['Customer']

    # Standardize the Phone Number
    df['Phone Number'] = df['Phone Number'].replace(regex=r'[^0-9.]', value='')
    df['Phone Number'] = '(' + df['Phone Number'].str[:3] + ') ' + df['Phone Number'].str[3:6] + '-' + df['Phone Number'].str[6:]
    # df['Phone Number']

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
                    'Customer ID',
                    'Lob Exclusion',
                    'Banking KWs',
                    'Life KWs',
                    'Spanish KWs',
    #                 'State',
                    'State_long',
                    'Phone Number',
                    'Location',
                    'Associate ID'
        ]]


    ## Spell checking LOB Exclusions

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

        ## Sitelink Upload --------------

    Sitelink_df = Campaign_df.copy()
    ref = pd.read_excel(ref_table, sheet_name='Sitelink') # read Ref table

    def Sitelink_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-00-7700-00000', row['Account Number'], regex=True, inplace=True)
            data_for.replace('https://www.URL.com/localagent', row['Website'], regex=True, inplace=True)
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Sitelink_df = Sitelink_update(Sitelink_df)

    # # path to save 
    # path = f'{main_path}\Output\Sitelink Upload'

    # # Checks if path exists
    # if os.path.isdir(main_path) == False:
    #     os.mkdir(main_path)

    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Sitelink_df.to_excel(f'{path}\\New Agent Sitelink Upload - {today_date}.xlsx', 
    #                     sheet_name='Sitelink Upload', 
    #                     index=False)


        ## Snippet Upload --------------


    Snippet_df = Campaign_df.copy()
    ref = pd.read_excel(ref_table, sheet_name='Snippet') # read Ref table

    def Snippet_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-00-7700-00000', row['Account Number'], regex=True, inplace=True)
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Snippet_df = Snippet_update(Snippet_df)

    # # path to save 
    # path = f'{main_path}\Output\Snippet Upload'

    # # Checks if path exists
    # if os.path.isdir(main_path) == False:
    #     os.mkdir(main_path)

    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Snippet_df.to_excel(f'{path}\\New Agent Snippet Upload - {today_date}.xlsx', 
    #                     sheet_name='Structured Snippet Upload', 
    #                     index=False)


        ## Call Upload --------------
        
    Call_df = Campaign_df.copy()
    ref = pd.read_excel(ref_table, sheet_name='Call Template') # read ref table


    def Call_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-00-7700-00000', row['Account Number'], regex=True, inplace=True)
            data_for['Phone number'] = row['Phone Number']
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        return res

    Call_df = Call_update(Call_df)

    # # path to save 
    # path = f'{main_path}\Output\Call Upload'

    # # Checks if path exists
    # if os.path.isdir(main_path) == False:
    #     os.mkdir(main_path)

    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # Call_df.to_excel(f'{path}\\New Agent Call Upload - {today_date}.xlsx', 
    #                     sheet_name='Call Upload', 
    #                     index=False)


        ## Location Upload -------------------------
    # - Agents with multiple Geo Target are excluded
    # - Agents without mileage are excluded

    radius_df = Campaign_df[['Account','Customer','Account Number','Geo Target']].copy()
    ref_1 = pd.read_excel(ref_table, sheet_name='Geo code Coordinates', skiprows=[0,1]) # read Ref table - SF
    ref_2 = pd.read_excel(ref_table, sheet_name = 'Campaign-Template')


    # Campaign Name Creation
    def Campaign_Name(x):
        res = ref_2['Campaign'][0].replace('[agentname]', x['Customer'])
        res = res.replace('000-00-7700-00000', x['Account Number'])
        return res
    radius_df['Campaign'] = radius_df.apply(Campaign_Name, axis=1)


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
    #     print('5',zip_code_list)
    #     print('6',miles_list)    
        
        
        # Create the rows for each element
        for i in range(len(miles_list)):
            zip_code = zip_code_list[i]
            mile = miles_list[i]
            
            res.loc[len(res.index)] = [x['Account'], x['Campaign'], zip_code, '', '', '', '']
            
            lat_long = ref_1[ref_1['Zip'] == int(zip_code)]['Lat_long'].values
            if len(lat_long) > 0:
                lat, long = lat_long[0].split(':')
    #             print('4', lat, long)
    #             print([x['Account'], x['Campaign'], None, lat, long, mile, 'mi'])
                if mile != '':
                    res.loc[len(res.index)] = [x['Account'], x['Campaign'], None, lat, long, mile, 'mi']
    #             print(res)
    #         print([x['Account'], x['Campaign'], zip_code[0], '', '', '', ''])      
        
    #     print('----------------')


    radius_df = res.copy()

    radius_df['Row Type'] = radius_df['Location'].apply(lambda x: 'proximity target' if x == None else 'location target')
    radius_df['Action'] = 'create'

    radius_df = radius_df[['Row Type', 'Action', 'Account', 'Campaign', 'Location', 'Latitude', 'Longitude', 'Location Radius','Location Radius Units']]

    # # path to save 
    # path = f'{main_path}\Output\Location Upload'

    # # Checks if path exists
    # if os.path.isdir(main_path) == False:
    #     os.mkdir(main_path)

    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # radius_df.to_excel(f'{path}\\New Agent Radius Target Upload - {today_date}.xlsx', 
    #                     sheet_name='Radius Location Upload', 
    #                     index=False)


        ### BULK UPLOAD -------------------
        ## 1) Campaign Type ---------------

    Campaign_Type_df = Campaign_df.copy()

    # Campaign Name Creation
    def Campaign_Name(x):
        res = ref_2['Campaign'][0].replace('[agentname]', x['Customer'])
        res = res.replace('000-00-7700-00000', x['Account Number'])
        return res
    Campaign_Type_df['Campaign'] = Campaign_Type_df.apply(Campaign_Name, axis=1)

    # Remove the State from Exclusion list
    Campaign_Type_df['State_long'] = Campaign_Type_df['State_long'].apply(lambda x: x + ', United States')

    exclusion_list = ref_2['Exclusion'][0]
    Campaign_Type_df['Exclusion'] = Campaign_Type_df['State_long'].apply(lambda x: exclusion_list.replace(x+'; ', ''))


    def remove_duplicates_list(x):
        return list(dict.fromkeys(x))


    # Main Function to create all columns for Campaign Type
    def Campaign_row_type(x):
        x['Row Type'] = 'Campaign'
        x['Action'] = 'Add'
        x['Campaign status'] = 'Enabled'

    #     x['Campaign start date (YYYY-MM-DD)'] = date.today().strftime("%Y-%m-%d")
    #     year_ahead = (date.today() + relativedelta(months=+x['Duration']))
    #     x['Campaign end date (YYYY-MM-DD)'] = year_ahead.replace(day=pd.Period(str(year_ahead)).days_in_month).strftime("%Y-%m-%d")
        x['Networks'] = 'Bing Search'
        x['Budget'] = 5
        x['Bid strategy type'] = 'Manual CPC'
        x['Delivery method'] = 'Standard'
        x['Language'] = 'EN'
        x['Label'] = x['State_long'].replace(', United States', '')
    #     to_remove = x['State']
    #     if to_remove == 'AK':
    #         x['Excluded region'] = Excluded_region[8:]
    #     else :
    #         x['Excluded region'] = Excluded_region.replace(f' | US-{to_remove}', '')

    #     x['Excluded country'] = Excluded_country
        x['Custom Parameter'] = '{_sfassociateid} = ' + x['Associate ID']
        x['Metro'] = ''

        return x[['Row Type', 'Action', 'Campaign status', 'Customer ID', 'Campaign', 'Networks',
                'Budget', 'Bid strategy type', 'Delivery method',
                'Language', 'Exclusion', 'Label', 'Custom Parameter']]
                

    Bulk_df_Campaign = Campaign_Type_df.apply(Campaign_row_type, axis=1)


        ## 2) Ad Group Type ---------------

    ref = pd.read_excel(ref_table, sheet_name='AdGroup-Template') # read ref table
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
        Lob_Exclusions = ['Financial' if KW == 'Financial Services' else KW for KW in Lob_Exclusions]
        
        if 'Mobile Home' in Lob_Exclusions:
            Lob_Exclusions.extend(['Mobile Home', 'Trailer'])
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
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-00-7700-00000', row['Account Number'], regex=True, inplace=True)
            
            data_for['Campaign start date'] = date.today().strftime("%Y-%m-%d")
            year_ahead = (date.today() + relativedelta(months=+row['Duration']))
            data_for['Campaign end date'] = year_ahead.replace(day=pd.Period(str(year_ahead)).days_in_month).strftime("%Y-%m-%d")
            
            data_for['Action'] = 'Add'
            data_for['Ad group status'] = 'Enabled'
            i = 0
            for row2 in data_for['Ad Group']:
                if row2.split(' - ')[2] in row['Lob Exclusion']:
    #                 print(row2.split(' - ')[1], row['Lob Exclusion'])
                    data_for['Ad group status'][i] = 'Paused'
    #                 data_for['Action'][i] = 'Add|Pause'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            data_for['Customer ID'] = row['Customer ID']
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
        
        res['Row Type'] = 'Ad Group'
        res['Default max. CPC'] = 5
        res['Language'] = 'EN'
        res['Ad rotation'] = 'Optimize'
        return res

    Bulk_df_AdGroup = AdGroup_update(AdGroup_df)

        ## 3) Ad Type ---------------

    ref = pd.read_excel(ref_table, sheet_name='Ad-Template') # read ref table
    Ad_df = Campaign_df.copy()

    Ad_df['Customer First Name'] = Ad_df['Customer'].str.split(' ',expand=True)[0]
    Ad_df['Lob Exclusion'] = Ad_df.apply(LOB, axis=1)

    def Ad_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()

            # Action and Ad Group Column Edit        
            data_for['Action'] = 'Add'
            data_for['Ad status'] = 'Enabled'
            i = 0
            for row2 in data_for['Ad Group']:
                if row2.split(' - ')[2] in row['Lob Exclusion']:
    #                 print(row2.split(' - ')[1], row['Lob Exclusion'])
    #                 data_for['Action'][i] = 'Create|Pause'
                    data_for['Ad status'][i] = 'Paused'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            
            # Ad Description and Ad Title Columns Lenght Edit
            for column in ['Description', 'Description 2', 'Description 3', 'Description 4']:
                i = 0
                for row3 in data_for[column]:
                    if len(row3) + len(row['Customer']) - agentname_len > 90:
                        data_for[column][i] = data_for[column][i].replace('agentname', row['Customer First Name'])
    #                     print(len(data_for[column][i]))
    #                     if len(data_for[column][i]) > 90:
    #                         print(row['Customer'])
    #                         print(data_for[column][i])
                    else: 
                        data_for[column][i] = data_for[column][i].replace('agentname', row['Customer'])
                    i += 1
                
            for column in ['Headline 1','Headline 2','Headline 3','Headline 4','Headline 5','Headline 6',
                        'Headline 7','Headline 8','Headline 9','Headline 10','Headline 11','Headline 12',
                        'Headline 13','Headline 14','Headline 15']:
                i = 0
                for row3 in data_for[column]:
    #                 print('Agent Name:', row['Customer'])
                    
                    if len(row3) + len(row['Customer']) - agentname_len > 30:
                        data_for[column][i] = data_for[column][i].replace('agentname', row['Customer First Name'])
                        
    #                     if len(data_for[column][i]) > 30:
    #                         print(row['Customer'])
    #                         print(data_for[column][i])
                    else: 
                        data_for[column][i] = data_for[column][i].replace('agentname', row['Customer'])
                    i += 1
        

            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-00-7700-00000', row['Account Number'], regex=True, inplace=True)
            data_for['Final URL'].replace('https://www.URL.com/localagent', row['Website'], regex=True, inplace=True)
            data_for['Customer ID'] = row['Customer ID']
            res = pd.concat([res,data_for], axis=0, ignore_index=True)
            
            
        res['Row Type'] = 'Ad'
        return res


    agentname_len = len('agentname')

    Bulk_df_Ad = Ad_update(Ad_df)


        ## 4) Keyword Type ---------------


    ref = pd.read_excel(ref_table, sheet_name='Keyword-Template') # read ref table
    Keyword_df = Campaign_df.copy()

    Keyword_df['Customer First Name'] = Keyword_df['Customer'].str.split(' ',expand=True)[0]
    Keyword_df['Lob Exclusion'] = Keyword_df.apply(LOB, axis=1)
    
    def Keyword_update(data):
        res = pd.DataFrame(columns=list(ref.columns))
        for i, row in data.iterrows():
            data_for = ref.copy()
            
    # Action and Ad Group Column Edit        
            data_for['Action'] = 'Add'
            data_for['Keyword status'] = 'Enabled'
            i = 0
            for row2 in data_for['Ad Group']:
                
                if row2.split(' - ')[2] in row['Lob Exclusion']:
                    data_for['Keyword status'][i] = 'Paused'

    #                 data_for['Action'][i] = 'Create|Pause'
                    data_for['Ad Group'][i] = 'INACTIVE LOB - ' + data_for['Ad Group'][i]
                i += 1
            
            
            data_for.replace('\[agentname]', row['Customer'], regex=True, inplace=True)
            data_for.replace('000-00-7700-00000', row['Account Number'], regex=True, inplace=True)
            
            data_for['Customer ID'] = row['Customer ID']
            res = pd.concat([res,data_for], axis=0, ignore_index=True)

        res['Row Type'] = 'Keyword'
        res.rename(columns={'Max CPC':'Default max. CPC',
                            'Keyword match type':'Type'}, inplace=True)
        return res


    Bulk_df_Keyword = Keyword_update(Keyword_df)

        ## 5) Merge four Subtable to One -----------------

    columns = list(Bulk_df_Campaign.columns) + list(Bulk_df_AdGroup.columns) + list(Bulk_df_Ad.columns) + list(Bulk_df_Keyword.columns)
    columns = list(dict.fromkeys(columns))
    # columns

    order_of_columns = 'Row Type	Action	Campaign status	Customer ID	Campaign	Networks	Budget	Bid strategy type	Delivery method	Campaign start date	Campaign end date	Language	Location	Exclusion	Label	Custom Parameter	Ad rotation	Ad group status	Ad Group	Default max. CPC	Keyword status	Keyword	Type	Ad status	Ad type	Headline 1	Headline 2	Headline 3	Headline 4	Headline 5	Headline 6	Headline 7	Headline 8	Headline 9	Headline 10	Headline 11	Headline 12	Headline 13	Headline 14	Headline 15	Description	Description 2	Description 3	Description 4	Path 1	Path 2	Final URL'
    order_of_columns = order_of_columns.split('\t')

    Merged_df = pd.DataFrame(columns=order_of_columns)

    bulk_df = pd.concat([Merged_df, Bulk_df_Campaign, Bulk_df_AdGroup, Bulk_df_Ad, Bulk_df_Keyword]).reset_index(drop=True)

    # # path to save 
    # path = f'{main_path}\Output'

    # # Checks if path exists
    # if os.path.isdir(path) == False:
    #     os.mkdir(path)
        
    # # Creates the doc
    # bulk_df.to_csv(f'{path}\\New Agent Bulk Upload - {today_date}.csv', 
    # #                     sheet_name='DS Upload Sheet', 
    #                     index=False)

    return Sitelink_df, Snippet_df, Call_df, radius_df, bulk_df