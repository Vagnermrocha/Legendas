# Corretor de legendas SRT

Aplicativo em Python com interface Tkinter para limpar arquivos SRT e traduzir legendas.

## Instalação

```powershell
python -m pip install -r requirements.txt
python corretor_srt.py
```

## Tradução

- **Google (online)** é o padrão e usa português brasileiro (`pt-BR`). Requer internet e o serviço público pode impor limites temporários de requisições.
- **Argos (offline)** funciona sem chamadas de tradução após instalar o modelo. O modelo disponível é inglês → português genérico e pode usar vocabulário de Portugal; o modelo é baixado na primeira utilização.

O aplicativo também remove formatação, propaganda e descrições entre colchetes, como `[woman screams]`.
