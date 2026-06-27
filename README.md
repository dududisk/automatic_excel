# Automação de Validação de Usuários (PyAutoGUI)

Automatiza o login de vários usuários de uma planilha Excel em um sistema web,
gerando um relatório final com o resultado de cada tentativa.

## Instalação

```bash
pip install -r requirements.txt
```

### OCR (detecção de "Senha incorreta" / "Usuário inválido")

O robô lê a tela com OCR para classificar o resultado de cada login. Para isso
é necessário o **motor Tesseract OCR** instalado no sistema:

1. Instale o Tesseract (Windows):

   ```powershell
   winget install --id UB-Mannheim.TesseractOCR -e
   ```

   Caminho padrão esperado: `C:\Program Files\Tesseract-OCR\tesseract.exe`
   (ajustável em `CAMINHOS_TESSERACT` no `main.py`).

2. O idioma português (`por.traineddata`) já vem na pasta `tessdata/` do
   projeto — o `main.py` a usa automaticamente via `TESSDATA_PREFIX`, sem
   precisar de permissão de administrador.

> Sem o Tesseract, a automação **continua funcionando**, mas todo login sem
> exceção é tratado como "Login realizado com sucesso" (sem distinguir senha
> incorreta de usuário inválido).
>
> ⚠️ Ajuste as listas `TEXTOS_SENHA_INCORRETA` e `TEXTOS_USUARIO_INVALIDO` no
> `main.py` para corresponderem às mensagens reais exibidas pelo site.

## Antes de executar

1. **Calibre as coordenadas.** Abra o site manualmente, posicione o mouse
   sobre cada elemento (campo login, senha, botão entrar, menu, sair) e use:

   ```bash
   python -c "import pyautogui, time; time.sleep(3); print(pyautogui.position())"
   ```

   Copie os valores para o bloco `# COORDENADAS` no início de `main.py`.

2. **Configure os caminhos e a URL.** Por padrão são lidos de variáveis de
   ambiente (recomendado, para não versionar dados pessoais):

   ```powershell
   $env:PLANILHA_DIR  = "C:\caminho\para\cadastrohes"
   $env:RELATORIO_DIR = "C:\caminho\para\relatorios"
   $env:SITE_URL      = "https://seu-sistema-de-login"
   ```

   Alternativamente, edite os valores padrão no bloco de configuração de
   `main.py` (`PASTA_PLANILHA`, `PASTA_RELATORIO`, `URL_SITE`).

3. **Ajuste os tempos** (`TEMPO_*`) conforme a velocidade da sua máquina/rede.

## Execução

```bash
python main.py
```

## Segurança / Parada de emergência

Mova o mouse para o **canto superior esquerdo** da tela a qualquer momento
para interromper imediatamente a automação (FailSafe do PyAutoGUI).

## Saída

Relatório `resultado_validacao.xlsx` na pasta definida em `RELATORIO_DIR`
(ou `PASTA_RELATORIO`), com as colunas: **Número, Nome, Login, Status**.
p
Status possíveis: `Login realizado com sucesso`, `Senha incorreta`,
`Usuário inválido`, `Erro inesperado`, `Interrompido pelo usuário`.
