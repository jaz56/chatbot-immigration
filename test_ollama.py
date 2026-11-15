import os
from dotenv import load_dotenv
from ollama import Client

load_dotenv()

client = Client(
    host="https://ollama.com",
    headers={"Authorization": f"Bearer {os.environ['OLLAMA_API_KEY']}"}
)

response = client.chat(
    model="gemma4:31b-cloud",
    messages=[{"role": "user", "content": "Bonjour, réponds en une phrase"}]
)

print(response.message.content)