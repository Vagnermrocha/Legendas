# Corretor de legendas e conversor Word para PDF

Aplicativo em Python com interface Tkinter para limpar e traduzir arquivos SRT e converter documentos Word em PDF preservando sua formatação.

## Instalação

```powershell
python -m pip install -r requirements.txt
python corretor_srt.py
```

## Conversão de Word para PDF

Use **Abrir Word...** na aba do corretor ou vá para a aba **Word para PDF** e clique em **Selecionar arquivo Word...**. Selecione um arquivo `.doc` ou `.docx` e escolha onde salvar o PDF. A conversão usa o LibreOffice, se instalado; caso contrário, usa o Microsoft Word por meio do pacote `docx2pdf`. Para essa segunda opção, é necessário ter o Microsoft Word instalado e ativado.

## Tradução

- **Google (online)** é o padrão e usa português brasileiro (`pt-BR`). Requer internet e o serviço público pode impor limites temporários de requisições.
- **Argos (offline)** funciona sem chamadas de tradução após instalar o modelo. O modelo disponível é inglês → português genérico e pode usar vocabulário de Portugal; o modelo é baixado na primeira utilização.

O aplicativo também remove formatação, propaganda e descrições entre colchetes, como `[woman screams]`.
