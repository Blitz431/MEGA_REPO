
# This script provides a client for interacting with the Ollama API.
# It includes a function pull_text that takes a prompt, sends it to the Ollama API,
# and returns the generated text if the request is successful.

from requests import post

# Ollama API URL
OLLAMA_API_URL = "http://localhost:11434/api/generate"

# Model for the API request
MODEL = "qwen2.5:3b"

def pull_text(prompt):
    """Pull text from the Ollama API using the provided prompt."""
    headers = {
        "Content-Type": "application/json"
    }
    data = {
        "prompt": prompt,
        "model": MODEL,  
        "stream": False
    }
    response = post(OLLAMA_API_URL, headers=headers, json=data)
    if response.status_code == 200:
        return response.json()["response"]
    else:
        raise Exception(f"Error pulling text from Ollama API: {response.status_code} - {response.text}")
