from dotenv import load_dotenv
from motor.motor_asyncio import AsyncIOMotorClient
import os


load_dotenv()
client = AsyncIOMotorClient(os.getenv("MONGO_URL"))
db = client["penn_stainless"]
logs = db['logs']