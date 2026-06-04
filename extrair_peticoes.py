# -*- coding: utf-8 -*-
import pdfplumber
import os
import json

pasta = os.path.join(os.path.dirname(__file__), "dados")
resultado = {}

for nome in sorted(os.listdir(pasta)):
    if not nome.lower().endswith(".pdf"):
        continue
    caminho = os.path.join(pasta, nome)
    try:
        texto = ""
        with pdfplumber.open(caminho) as pdf:
            for pagina in pdf.pages[:3]:
                t = pagina.extract_text()
                if t:
                    texto += t + "\n"
        resultado[nome] = texto[:3000]
    except Exception as e:
        resultado[nome] = f"ERRO: {e}"

with open(os.path.join(os.path.dirname(__file__), "textos_peticoes.json"), "w", encoding="utf-8") as f:
    json.dump(resultado, f, ensure_ascii=False, indent=2)

print(f"Extraídos {len(resultado)} arquivos.")
