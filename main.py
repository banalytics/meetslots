import os
from gcsa.google_calendar import GoogleCalendar
from dotenv import load_dotenv

load_dotenv('.env')
calendar = GoogleCalendar(os.getenv('EMAIL'), credentials_path='credentials.json')

