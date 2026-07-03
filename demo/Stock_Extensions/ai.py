"""
PEAT EXTENSION
author = Beffy
name = PEAT Ollama Integration
filename = ai
version = 1.0

requirements:
- None
"""

# Requires oLlama to be installed and running on your computer

import requests

EXT_NAMESPACE = "ai"
ollama_link = "http://localhost:11434/api/generate"

help_dict = {
    "ask": "Ask a question to the local AI (Ollama) in arg1 (wrapped in quotes)",
}

def load_extension():
    peat.register_command("ask", cmd_ai_ask) # type: ignore

    peat.register_help(help_dict) # type: ignore

def ollama_ask(prompt, model="llama3"):
    try:
        r = requests.post(
            ollama_link,
            json={
                "model": model,
                "prompt": prompt,
                "stream": False
            }
        )
        return r.json().get("response", "").strip()

    except Exception as e:
        return f"[Ollama error] {e}"
    
def cmd_ai_ask(a1, a2, title):
    if not a1:
        peat.print("Expected a question.") # type: ignore
        return

    if a1[0] in ('"', "'"):
        a1 = peat.clean_args(a1, a1[0]) # type: ignore

    peat.print("Let me think...") # type: ignore

    reply = ollama_ask(a1)

    peat.print(f"Ollama says: {reply}") # type: ignore