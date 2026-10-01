import os

from dotenv import load_dotenv
from fastapi import FastAPI

load_dotenv()

app = FastAPI(title="Liquide Case API")


@app.get("/health")
def health():
    return {
        "status": "ok",
        "groq_key_loaded": bool(os.getenv("GROQ_API_KEY")),
    }