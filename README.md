# hours-tracking

### Purpose
I created this app to track my hours for a company I work for via a CLI. It will store start time, end time, total hours/minutes worked, reasons for work, and will create a table for you to send to whoever needs it.

This relies heavily on the excellent [gspread](https://gspread.readthedocs.io/en/latest/) library.


### Requirements
+ Python 3.6+
+ Google account


### Setup

1. Download this code.

2. Obtain [Oauth2 Google Credentials](https://console.developers.google.com/project), and save the JSON file somewhere memorable.

3. Create a Google Sheet where you want to store your hours.

4. Grant the API permission to access your Google Sheet ([tutorial](https://www.dundas.com/support/learning/documentation/connect-to-data/how-to/connecting-to-google-sheets))

5. Create the following envvars on your local machine:
```
  HOURS_SHEET_NAME=\<<your sheet name>>
  HOURS_JSON_CREDENTIALS_DIR=\<<directory of your JSON credentials>>
  HOURS_EMAIL_ADDRESS=\<<your gmail address>>
  HOURS_EMAIL_PASSWORD=\<<your gmail password>>
  ```
  
6. Install requirements.txt (```python pip install requirements.txt ```), setting up an venv if you prefer.

7. Change directory to the location of the code and run it! (``` python hours.py```).




