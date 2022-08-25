import tweepy
import webbrowser
import time
import pandas as pd

consumer_key = 'h1FMGukUb2opdWyp6e5h4wXsJ'
consumer_secret = 'Bd3604rdoBGdsVdBL6fubOpyYL44L0uJNpBp5D63IOe5k4h6Rg'
access_token = '359263757-cULZ2ajnBbST0hKg34oWkSKfyw5WHr3t0hAD1ghZ'
access_token_secret = 'vsmYdUvfGISnwpjftR2R0haRyV7SGPHTUsCVxoRirh4Mj'

auth = tweepy.OAuthHandler(consumer_key,consumer_secret)
auth.set_access_token(access_token,access_token_secret)

api = tweepy.API(auth)

def extract_discom_data(area_list):
    uppcl_data = []
    columns_headers = ['USER OR HANDLE','NAME','FOLLOWERS','TWEETS']
    for area in area_list:
        try:
            user = api.get_user(area)
        except:
            print(area)
            
        single_area_data = {'USER OR HANDLE': user.screen_name,'NAME':user.name,'FOLLOWERS':user.followers_count,'TWEETS':user.statuses_count}
        uppcl_data.append(single_area_data)

    data = pd.DataFrame(uppcl_data,columns=columns_headers)
    return data

df = pd.read_excel("./twitterhandles.xlsx",sheet_name="management")

result = extract_discom_data(df['handles'].tolist())
result.to_excel(r"C:\Users\DELL\Desktop\Dev\result.xlsx")
print(result['FOLLOWERS'].describe())