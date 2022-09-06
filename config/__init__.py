import os
from dotenv import load_dotenv
from pathlib import Path



# path = "./config/." + os.getenv("env") + ".env"
# load_dotenv(dotenv_path=path)

path = f"./config/.{os.environ.get('app_env')}.env"
load_dotenv(dotenv_path=path)

# print(f"Running config path: {path} {os.environ}")

# def load_os_dotenv() :
#     path = Path(f'{os.getcwd()}/config/.{os.environ.get("env")}.env')
#     load_dotenv(path)
#     if os.path.exists(path):
#         load_dotenv(path)
