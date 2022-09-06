import os
import psycopg2
# from config import load_os_dotenv

# print(os.getenv("POSTGRES_HOST"))
# load_os_dotenv()
main = psycopg2.connect(
    host=os.getenv("POSTGRES_HOST"),
    port=os.getenv("POSTGRES_PORT"),
    database=os.getenv("POSTGRES_DATABASE"),
    user=os.getenv("POSTGRES_USER"),
    password=os.getenv("POSTGRES_PASS")
)