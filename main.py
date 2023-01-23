import pandas as pd
import os
from dotenv import load_dotenv
from pathlib import Path
import pandas as pd
import time

# selenium 4
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

env_path = Path('.env')
load_dotenv(env_path)

login_url = os.getenv('URL')
wp_username = os.getenv('WP_USERNAME')
wp_password = os.getenv('WP_PASSWORD')
submissions_url = os.getenv('SUBMISSIONS_URL')
client_name = os.getenv('CLIENT_NAME')


# Headless Chrome
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    'download.default_directory': rf"{os.getcwd()}",
})
chrome_options.add_experimental_option("detach", True)
chrome_options.headless = True

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)

driver.get(login_url)

driver.implicitly_wait(1)
assert "Log In" in driver.title

username_field = driver.find_element(by=By.ID, value="user_login")
username_field.send_keys(wp_username)

password_field = driver.find_element(by=By.ID, value="user_pass")
password_field.send_keys(wp_password + Keys.ENTER)

driver.get(submissions_url)

driver.implicitly_wait(5)

export_button = driver.find_elements(by=By.CLASS_NAME, value="e-export-button")[0]
time.sleep(5)
export_button.click()

time.sleep(10)

driver.quit()

# loop through current folder and open the csv file
for file in os.listdir(os.getcwd()):
    if file.endswith(".csv"):
        df = pd.read_csv(file, header=0, index_col=False, usecols=["Form Name (ID)", "Created At", "Referrer", "Email", "Name"])


# convert created at to datetime
df['Created At'] = pd.to_datetime(df['Created At'])

# group by month and count
grouped_df = df.groupby([pd.Grouper(key='Created At', freq='M'), 'Form Name (ID)']).count()

# rename "Email" column to "Total"
grouped_df.rename(columns={'Email': 'Total'}, inplace=True)

# change month dates to name of month and year
grouped_df['Created At'] = grouped_df.index.get_level_values(0).strftime('%B %Y')

# rename "Created At" column to "Month"
grouped_df.rename(columns={'Created At': 'Month'}, inplace=True)

grouped_df = grouped_df[['Month', 'Total']]

# drop index
grouped_df.reset_index(drop=True, inplace=True)

# get current date
current_date = pd.to_datetime('today').strftime('%Y-%m-%d')

with pd.ExcelWriter(f'{client_name} Submissions {current_date}.xlsx') as writer:
    df.to_excel(writer, sheet_name='Submissions', index=False)
    grouped_df.to_excel(writer, sheet_name='Analysis', index=False)

