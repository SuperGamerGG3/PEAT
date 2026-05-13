"""
PEAT EXTENSION
author = Beffy
name = ai
version = 1.0

requirements:
- none
"""

# Requires oLlama to be installed and running on your computer

import requests

EXT_NAMESPACE = "ai"

help_dict = {
    "ask": "Ask a question to the local AI (Ollama) in arg1 (wrapped in quotes)",
}

def load_extension():
    peat.register_command(EXT_NAMESPACE, "ask", cmd_ai_ask) # type: ignore

    peat.register_help(EXT_NAMESPACE, help_dict) # type: ignore

def ollama_ask(prompt, model="llama3"):
    try:
        r = requests.post(
            "http://localhost:11434/api/generate",
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
        peat.voice_print("Expected a question.") # type: ignore
        return

    if a1[0] in ('"', "'"):
        a1 = peat.clean_args(a1, a1[0]) # type: ignore

    peat.voice_print("Thinking...") # type: ignore

    reply = ollama_ask(a1)

    peat.voice_print(reply) # type: ignore