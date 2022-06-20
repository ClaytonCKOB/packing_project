from pathlib import Path
import os

BASE_DIR = Path(__file__).resolve().parent.parent
PASSWORD = os.environ['PASSWORD']
DB_PATH = os.environ['DB_PATH']
DB_CONFIG_PATH = os.environ['DB_CONFIG_PATH']
MAIN_EMAIL = os.environ['MAIN_EMAIL'] 