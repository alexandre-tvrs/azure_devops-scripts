from os import environ as env
from dotenv import load_dotenv

load_dotenv()

global organization
global token

organization = env['ORGANIZATION']
token = env['TOKEN']
