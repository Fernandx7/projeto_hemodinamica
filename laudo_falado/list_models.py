import google.generativeai as genai
import os

CHAVE_API_PATH = "chaveapi"

try:
    with open(CHAVE_API_PATH, "r") as f:
        api_key = f.read().strip()
    genai.configure(api_key=api_key)
    
    print("Available models:")
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            print(f"- {m.name}")
except Exception as e:
    print(f"Error: {e}")
