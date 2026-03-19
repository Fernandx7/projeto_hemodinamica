import google.generativeai as genai

# Configura sua chave
genai.configure(api_key="AIzaSyBOFWPE2HnK3LXjMWJuNUr2R7SAeKveP-o")

# 1. Faz o upload do arquivo para a API (ele fica guardado por 48h)
audio_file = genai.upload_file(path="")

# 2. Chama o modelo pedindo a transcrição
model = genai.GenerativeModel('gemini-1.5-flash')
response = model.generate_content([
    "Transcreva este áudio detalhadamente. Se houver termos técnicos médicos, certifique-se de grafá-los corretamente.",
    audio_file
])

print(response.text)
