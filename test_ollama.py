import ollama

response = ollama.chat(
    model="llama3",
    messages=[{"role": "user", "content": "한 문장으로만 답해. 잘 연결됐어?"}],
)
print(response["message"]["content"])