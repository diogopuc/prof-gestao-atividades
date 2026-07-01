"""
Automação - Agenda Docente PUCPR
Extrai nomes de professores da página de Aprovações > Gestão de Atividades
e gera um arquivo Excel com os resultados.

Dependências:
    pip install playwright openpyxl
    playwright install msedge
"""

from playwright.sync_api import sync_playwright
import openpyxl
import math
import os
import time

# ============================================================
#  CONFIGURAÇÃO — ajuste aqui antes de rodar
# ============================================================
URL_LOGIN     = "https://agendadocente.pucpr.br"
DATA_INICIO   = "01/07/2026"   # DD/MM/AAAA  ← altere aqui
DATA_FIM      = "07/07/2026"   # DD/MM/AAAA  ← altere aqui
ARQUIVO_SAIDA = f"professores_{DATA_INICIO.replace('/','_')}_{DATA_FIM.replace('/','_')}.xlsx"

# Perfil padrão do Edge (mesmo usado no dia a dia, com sessões/senhas salvas).
# IMPORTANTE: feche todas as janelas do Edge antes de rodar o script,
# pois o Windows bloqueia o perfil enquanto ele estiver em uso.
EDGE_USER_DATA_DIR = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\User Data")
# ============================================================


def selecionar_intervalo_datas(page, data_inicio: str, data_fim: str):
    """
    Abre o calendário e seleciona o intervalo clicando nas duas datas.
    O calendário funciona com dois cliques: primeiro = início, segundo = fim.
    """
    dia_ini  = int(data_inicio.split("/")[0])
    mes_ini  = int(data_inicio.split("/")[1])
    ano_ini  = int(data_inicio.split("/")[2])
    dia_fim  = int(data_fim.split("/")[0])
    mes_fim  = int(data_fim.split("/")[1])
    ano_fim  = int(data_fim.split("/")[2])

    MESES_PT = {
        "Janeiro":1,"Fevereiro":2,"Março":3,"Abril":4,
        "Maio":5,"Junho":6,"Julho":7,"Agosto":8,
        "Setembro":9,"Outubro":10,"Novembro":11,"Dezembro":12
    }

    def abrir_calendario():
        # Clica no campo de data (ícone de calendário ou input)
        page.locator(".sc-aXZVg.sc-gEvEer").first.click()
        time.sleep(0.5)

    def mes_ano_atual():
        texto = page.locator("text=/^(Janeiro|Fevereiro|Março|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro)\\s+\\d{4}$/").inner_text(timeout=5000)
        partes = texto.strip().split()
        return MESES_PT[partes[0]], int(partes[1])

    def navegar_ate(mes_alvo, ano_alvo):
        for _ in range(36):
            m, a = mes_ano_atual()
            if m == mes_alvo and a == ano_alvo:
                break
            if (a, m) < (ano_alvo, mes_alvo):
                # clicar seta direita (próximo mês)
                page.locator("text=/Julho|Janeiro|Fevereiro|Março|Abril|Maio|Junho|Agosto|Setembro|Outubro|Novembro|Dezembro/").locator("..").locator("svg").last.click()
            else:
                page.locator("text=/Julho|Janeiro|Fevereiro|Março|Abril|Maio|Junho|Agosto|Setembro|Outubro|Novembro|Dezembro/").locator("..").locator("svg").first.click()
            time.sleep(0.3)

    def clicar_dia(dia: int):
        # Clica na célula do dia no calendário
        page.locator(f"text={dia}").filter(has_text=f"^{dia}$").or_(
            page.get_by_role("cell", name=str(dia), exact=True)
        ).first.click()
        time.sleep(0.3)

    abrir_calendario()
    navegar_ate(mes_ini, ano_ini)
    clicar_dia(dia_ini)
    navegar_ate(mes_fim, ano_fim)
    clicar_dia(dia_fim)

    # Fecha o calendário pressionando Escape
    page.keyboard.press("Escape")
    time.sleep(0.3)


def extrair_nomes(page) -> list:
    """Executa o mesmo JS do processo manual para extrair nomes da página atual."""
    return page.evaluate("""
        () => {
            return Array.from(document.querySelectorAll('span'))
                .map(e => e.innerText.trim())
                .filter(texto =>
                    /^[A-ZÁÉÍÓÚÃÕÇ\\s]+$/.test(texto) &&
                    texto.length > 3
                );
        }
    """)


def obter_total_e_por_pagina(page) -> tuple:
    """Lê '1-10 de 219' e retorna (total_registros, itens_por_pagina, total_paginas)."""
    try:
        els = page.locator(".sc-laNGHT.lgZpOK").all()
        for el in els:
            txt = el.inner_text()
            if "de" in txt and "-" in txt:
                # Formato: "1-10 de 219"
                por_pagina = int(txt.split("-")[1].split(" ")[0].strip())
                total      = int(txt.split("de")[-1].strip())
                paginas    = math.ceil(total / por_pagina)
                print(f"  → Total de registros: {total} | Por página: {por_pagina} | Páginas: {paginas}")
                return total, por_pagina, paginas
    except Exception as e:
        print(f"  ⚠ Não foi possível ler paginação: {e}")
    return 0, 10, 1


def clicar_proxima_pagina(page) -> bool:
    """
    Clica no segundo SVG da div de paginação (seta direita = próxima página).
    Retorna False se o botão estiver desabilitado ou não encontrado.
    """
    try:
        paginacao = page.locator(".sc-cHMHOW.brQzqN").first
        svgs = paginacao.locator("svg").all()
        if len(svgs) >= 2:
            proximo = svgs[1]  # índice 1 = seta direita
            proximo.click()
            time.sleep(1.5)  # aguarda carregar a próxima página
            return True
    except Exception as e:
        print(f"  ⚠ Erro ao clicar próxima página: {e}")
    return False


def salvar_excel(nomes: list, arquivo: str):
    """Salva a lista de nomes em um arquivo Excel, um nome por linha."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Professores"

    # Cabeçalho
    ws.append(["#", "Nome do Professor"])

    for i, nome in enumerate(nomes, start=1):
        ws.append([i, nome])

    # Ajustar largura da coluna de nomes
    ws.column_dimensions["B"].width = 50

    wb.save(arquivo)
    print(f"\n✅ Excel salvo: {arquivo} ({len(nomes)} professores)")


def main():
    with sync_playwright() as p:
        # Usa o perfil padrão do Edge (mesmo do dia a dia, com sessões/senhas
        # salvas), em vez de abrir uma janela isolada tipo "anônima".
        # headless=False para ver o navegador funcionando.
        context = p.chromium.launch_persistent_context(
            EDGE_USER_DATA_DIR,
            channel="msedge",
            headless=False,
            slow_mo=200,
        )
        page = context.new_page()

        print("=" * 60)
        print("  Automação - Agenda Docente PUCPR")
        print(f"  Período: {DATA_INICIO} → {DATA_FIM}")
        print("=" * 60)

        # ── 1. Login ────────────────────────────────────────────────
        print("\n[1] Acessando a página de login...")
        page.goto(URL_LOGIN, wait_until="networkidle")
        input("\n>> Faça o login manualmente na janela do navegador e pressione ENTER aqui para continuar...\n")

        # ── 2. Acessar a página de workflows ───────────────────────
        print("[2] Acessando a página de Gestão de Atividades...")
        page.goto(URL_LOGIN + "/workflows", wait_until="networkidle")
        time.sleep(2)

        # ── 3. Selecionar "Gestão de Atividades" ──────────────────
        print("[3] Selecionando 'Gestão de Atividades'...")
        page.select_option("select", value="3")
        time.sleep(1.5)

        # ── 4. "Em andamento" já vem selecionado por padrão ───────
        print("[4] Status 'Em andamento' já está selecionado.")

        # ── 5. Selecionar o intervalo de datas ────────────────────
        print(f"[5] Selecionando datas: {DATA_INICIO} → {DATA_FIM}...")
        try:
            selecionar_intervalo_datas(page, DATA_INICIO, DATA_FIM)
        except Exception:
            page.screenshot(path="erro_calendario.png", full_page=True)
            print("\n⚠ Falha ao selecionar as datas. Screenshot salvo em 'erro_calendario.png' para diagnóstico.")
            raise
        time.sleep(1)

        # ── 6. Clicar em Filtrar ───────────────────────────────────
        print("[6] Clicando em Filtrar...")
        page.get_by_role("button", name="Filtrar").click()
        time.sleep(2)

        # ── 7. Extrair nomes percorrendo todas as páginas ──────────
        print("[7] Iniciando extração...")
        _, _, total_paginas = obter_total_e_por_pagina(page)

        todos_nomes = []

        for pagina_atual in range(1, total_paginas + 1):
            print(f"  📄 Página {pagina_atual}/{total_paginas}...", end=" ")
            nomes = extrair_nomes(page)
            todos_nomes.extend(nomes)
            print(f"{len(nomes)} nomes extraídos | Total acumulado: {len(todos_nomes)}")

            if pagina_atual < total_paginas:
                sucesso = clicar_proxima_pagina(page)
                if not sucesso:
                    print("  ⚠ Não foi possível avançar de página. Encerrando.")
                    break

        # ── 8. Remover duplicatas mantendo a ordem ─────────────────
        vistos = set()
        nomes_unicos = []
        for nome in todos_nomes:
            if nome not in vistos:
                vistos.add(nome)
                nomes_unicos.append(nome)

        print(f"\n  Total extraído: {len(todos_nomes)} | Únicos: {len(nomes_unicos)}")

        # ── 9. Salvar no Excel ─────────────────────────────────────
        salvar_excel(nomes_unicos, ARQUIVO_SAIDA)

        context.close()
        print("\n✅ Automação concluída com sucesso!")


if __name__ == "__main__":
    main()