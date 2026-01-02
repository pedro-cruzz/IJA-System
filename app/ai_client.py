import os
from openai import OpenAI

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def responder_ia(pergunta: str, contexto: str = "") -> str:
    resp = client.responses.create(
        model="gpt-5.2",
        instructions=(
            "Você é um assistente do painel ADMIN do sistema IJA. "
            "Responda em PT-BR, direto ao ponto, com passos práticos. "
            "Se faltar informação, peça 1 detalhe específico. "
            "Não invente dados do sistema."
        ),
        input=f"{contexto}\n\nPergunta:\n{pergunta}",
    )
    return (resp.output_text or "").strip()
