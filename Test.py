import pandas as pd

import Google.Google_Builder as GB

df = pd.read_excel('Data - SF.xlsx', sheet_name='Data')
Ref_df = 'Ref table - SF.xlsx'

Snippet_df, Sitelink_df, radius_df = GB.main(df, Ref_df)

import zipfile

zip_file = zipfile.ZipFile('file_name.zip', 'w')
zip_file.write(Snippet_df.to_excel('file.xlsx'))
    # '/tmp/hello.txt')

zip_file.close()