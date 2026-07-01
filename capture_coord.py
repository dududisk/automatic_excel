"""
Captura de coordenadas X/Y da tela para a automação (main.py).

Como usar:
    1. Abra o Chrome no site de login e deixe a janela do jeito que será usada
       na automação (mesmo tamanho, posição e zoom — isso é importante, pois
       as coordenadas dependem disso).
    2. Rode este script:  python capture_coord.py
    3. Para cada item pedido, posicione o mouse EM CIMA do elemento e, SEM
       mover, volte ao terminal e pressione ENTER (ou aguarde a contagem, no
       modo automático). O script grava o X/Y daquele ponto.
    4. No final, ele imprime um bloco pronto para colar no main.py e também
       salva em "coordenadas_capturadas.txt".

Dica: alguns itens (menu do usuário e botão sair) só aparecem DEPOIS do login.
Você pode rodar o script em duas etapas, ou logar manualmente uma vez para
deixar esses elementos visíveis enquanto captura.
"""

import os
import time

import pyautogui

# Lista dos pontos a capturar: (NOME_DA_VARIAVEL, descrição amigável).
# A ordem segue a sequência da automação no main.py.
PONTOS = [
    ("LOGIN",         "campo de LOGIN (usuário)"),
    ("SENHA",         "campo de SENHA"),
    ("BOTAO_ENTRAR",  "botão ENTRAR"),
    ("MENU_USUARIO",  "menu do usuário (aparece após o login)"),
    ("BOTAO_SAIR",    "botão SAIR / logout (aparece após o login)"),
]

# Segundos de contagem regressiva no modo automático (sem precisar do teclado).
SEGUNDOS_CONTAGEM = 5


def capturar_com_enter(descricao):
    """
    Captura a posição do mouse quando o usuário pressiona ENTER no terminal.

    Retorna uma tupla (x, y).
    """
    input(f"  -> Posicione o mouse sobre o {descricao} e pressione ENTER...")
    x, y = pyautogui.position()
    return x, y


def capturar_com_contagem(descricao, segundos=SEGUNDOS_CONTAGEM):
    """
    Captura a posição do mouse após uma contagem regressiva.

    Útil quando você não quer tirar a mão do mouse para apertar Enter.
    Retorna uma tupla (x, y).
    """
    print(f"  -> Leve o mouse até o {descricao}. Capturando em:")
    for restante in range(segundos, 0, -1):
        # Mostra a contagem e a posição atual do mouse em tempo real.
        x, y = pyautogui.position()
        print(f"     {restante}...  (mouse agora em X={x}, Y={y})   ", end="\r")
        time.sleep(1)
    x, y = pyautogui.position()
    print()  # quebra a linha do contador
    return x, y


def escolher_modo():
    """Pergunta qual modo de captura usar. Retorna 'enter' ou 'contagem'."""
    print("Como deseja capturar cada ponto?")
    print("  [1] Pressionando ENTER (recomendado)")
    print("  [2] Contagem regressiva automática (não precisa do teclado)")
    escolha = input("Escolha (1/2) [1]: ").strip()
    return "contagem" if escolha == "2" else "enter"


def formatar_bloco(coordenadas):
    """
    Monta o bloco de configuração no mesmo formato do main.py.

    `coordenadas` é um dict {NOME: (x, y)}.
    Retorna uma string pronta para colar.
    """
    linhas = []
    linhas.append("# === COORDENADAS CAPTURADAS - cole no bloco de config do main.py ===")
    for nome, _ in PONTOS:
        x, y = coordenadas[nome]
        linhas.append(f"{nome}_X = {x}")
        linhas.append(f"{nome}_Y = {y}")
        linhas.append("")  # linha em branco entre os pares
    return "\n".join(linhas).rstrip() + "\n"


def main():
    """Fluxo principal: guia a captura de todos os pontos e salva o resultado."""
    print("=" * 64)
    print(" CAPTURA DE COORDENADAS PARA A AUTOMACAO")
    print("=" * 64)
    print("Mantenha a janela do Chrome exatamente como sera na automacao.")
    print("(tamanho, posicao e zoom influenciam as coordenadas)\n")

    modo = escolher_modo()
    print()

    coordenadas = {}
    for nome, descricao in PONTOS:
        if modo == "contagem":
            x, y = capturar_com_contagem(descricao)
        else:
            x, y = capturar_com_enter(descricao)

        coordenadas[nome] = (x, y)
        print(f"  [ok] {nome}: X={x}, Y={y}\n")

    # Monta e exibe o bloco pronto.
    bloco = formatar_bloco(coordenadas)
    print("=" * 64)
    print(bloco)

    # Salva também em arquivo para consulta posterior (caminho absoluto ao
    # lado deste script, independente do diretório de trabalho atual).
    caminho_saida = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "coordenadas_capturadas.txt")
    with open(caminho_saida, "w", encoding="utf-8") as arquivo:
        arquivo.write(bloco)
    print(f"[INFO] Coordenadas salvas em: {caminho_saida}")
    print("[INFO] Copie os valores acima para o bloco de COORDENADAS do main.py.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[CANCELADO] Captura interrompida pelo usuario.")
