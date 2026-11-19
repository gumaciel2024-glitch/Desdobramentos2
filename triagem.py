import argparse
import re
import sys
import time
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import Error, sync_playwright

CNJ = r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}"
VAZIO = re.compile(r"nenhum(?:a)?\s+(?:registro|resultado|ordem|dado)[^\n]*|não foram encontrad[oa]s[^\n]*", re.I)
PAGINA = re.compile(r"\b(\d+)\s*[-–]\s*(\d+)\s+de\s+(\d+)\b", re.I)


def digitos(valor):
    return re.sub(r"\D", "", str(valor))


def classificar(linhas, total, cnj):
    if not linhas or (total is not None and total != len(linhas)):
        return None
    if any([digitos(n) for n in re.findall(CNJ, linha)] != [cnj] for linha in linhas):
        return None
    return "Ordem Exedente" if len(linhas) > 1 else ""


ICONES = """
row => [...row.querySelectorAll('td')].slice(0, 2).flatMap(cell =>
    [...cell.querySelectorAll('button, a, [role=button], mat-icon, i, svg')]
        .filter(el => el.getClientRects().length &&
                      !el.querySelector('mat-icon, i, svg'))
        .map(el => {
            const s = getComputedStyle(el);
            const p = el.querySelector('path');
            const b = el.closest('button, a, [role=button]');
            return JSON.stringify([
                el.tagName, el.textContent.trim(), s.color, s.fill, s.opacity,
                p && getComputedStyle(p).fill, b && getComputedStyle(b).opacity,
                b && (b.disabled === true || b.getAttribute('aria-disabled') === 'true'),
                b && getComputedStyle(b).backgroundColor
            ]);
        }))
"""


def linhas_tabela(pagina):
    tabela = pagina.get_by_role("table").filter(
        has=pagina.get_by_role("columnheader", name=re.compile("Número.*Protocolo", re.I))
    )
    return tabela.locator("tbody tr:visible").filter(has_text=re.compile(CNJ))


def consultar(pagina, numero, vazio, prazo=20):
    pagina.get_by_role("button", name="Limpar", exact=True).click()
    campo = pagina.get_by_role("textbox", name=re.compile("Número do Processo", re.I))
    campo.wait_for()
    linhas = linhas_tabela(pagina)
    if campo.input_value().strip(" ._-") or linhas.count():
        raise RuntimeError("A consulta anterior ainda aparece após Limpar.")
    campo.fill(numero)
    if digitos(campo.input_value()) != digitos(numero):
        raise RuntimeError("O campo não recebeu o número completo.")
    vazio_antes = pagina.get_by_text(vazio).filter(visible=True).count()
    campo.press("Enter")
    limite = time.monotonic() + prazo
    anterior = None
    while time.monotonic() < limite:
        textos = linhas.all_inner_texts()
        paginas = PAGINA.findall(pagina.locator("body").inner_text())
        total = int(paginas[0][2]) if len(paginas) == 1 else None
        sem_ordem = not vazio_antes and bool(pagina.get_by_text(vazio).filter(visible=True).count())
        estado = (textos, total, sem_ordem)
        if estado == anterior and textos:
            status = classificar(textos, total, digitos(numero))
            if status is not None:
                return status, linhas.first
        if not textos and sem_ordem and anterior == estado:
            return "Sem ordem", None
        anterior = estado
        time.sleep(0.4)
    raise RuntimeError("Resultado ausente, inesperado ou com paginação divergente.")


def numero_cnj(valor):
    texto = str(valor).strip()
    if re.fullmatch(CNJ, texto):
        return texto
    if re.fullmatch(r"\d{20}", texto):
        return f"{texto[:7]}-{texto[7:9]}.{texto[9:13]}.{texto[13]}.{texto[14:16]}.{texto[16:]}"
    raise ValueError("Número inválido. Guarde o CNJ completo como texto.")


def calibrar(pagina, vazio, rotulo):
    numero = numero_cnj(input(f"CNJ com uma única ordem e página {rotulo}: "))
    status, linha = consultar(pagina, numero, vazio)
    if status:
        raise RuntimeError(f"Exemplo de calibração: {status}")
    icones = linha.evaluate(ICONES)
    if len(icones) != 2:
        raise RuntimeError("A estrutura dos ícones da tabela precisa ser ajustada.")
    return icones[1]


def abrir_navegador(chromium):
    if sys.platform != "win32":
        return chromium.launch(headless=False)

    falhas = []
    for canal, nome in (("msedge", "Microsoft Edge"), ("chrome", "Google Chrome")):
        try:
            navegador = chromium.launch(channel=canal, headless=False)
        except Error as erro:
            falhas.append(f"{nome}: {str(erro).splitlines()[0]}")
            continue
        print(f"Navegador: {nome}")
        return navegador

    raise SystemExit(
        "Não foi possível abrir o Edge nem o Chrome instalados.\n"
        + "\n".join(falhas)
    )


def executar(args):
    caminho = Path(args.planilha)
    saida = caminho.with_name(caminho.stem + "_triagem.xlsx")
    if caminho.suffix.lower() != ".xlsx" or saida.exists():
        raise SystemExit("Use .xlsx e mantenha livre o nome da cópia _triagem.xlsx.")
    livro = load_workbook(caminho)
    aba = livro.active
    cabecalhos = {str(c.value).strip().casefold(): c.column for c in aba[1]}
    coluna_cnj = cabecalhos.get("número do processo")
    coluna_status = cabecalhos.get("status")
    if not coluna_cnj or not coluna_status:
        raise SystemExit("Use os cabeçalhos Número do Processo e Status na primeira linha.")
    vazio = re.compile(args.sem_ordem, re.I) if args.sem_ordem else VAZIO
    with sync_playwright() as p:
        navegador = abrir_navegador(p.chromium)
        pagina = navegador.new_page()
        pagina.goto(args.url)
        input("Entre no SISBAJUD, abra Ordens Judiciais > Busca por filtros de pesquisa e pressione Enter aqui. ")
        apagado = calibrar(pagina, vazio, "apagada")
        ativo = calibrar(pagina, vazio, "ativa")
        if apagado == ativo:
            raise RuntimeError("Os dois estados do ícone não puderam ser diferenciados.")
        livro.save(saida)
        for indice in range(2, aba.max_row + 1):
            destino = aba.cell(indice, coluna_status)
            entrada = aba.cell(indice, coluna_cnj).value
            if entrada is None or str(destino.value if destino.value is not None else "").strip():
                continue
            try:
                status, linha = consultar(pagina, numero_cnj(entrada), vazio)
                if not status:
                    icones = linha.evaluate(ICONES)
                    if len(icones) != 2 or icones[1] not in (apagado, ativo):
                        raise RuntimeError("Estado do ícone diferente da calibração.")
                    status = "NEGATIVO" if icones[1] == apagado else "Verificar operações"
            except Exception as erro:
                print(f"Linha {indice} para revisar: {str(erro).splitlines()[0]}")
                continue
            destino.value = status
            temporario = saida.with_suffix(".tmp.xlsx")
            livro.save(temporario)
            temporario.replace(saida)
            print(f"Linha {indice}: {status}")
        print(f"Cópia: {saida}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("planilha", nargs="?")
    parser.add_argument("url", nargs="?", help="Endereço do SISBAJUD")
    parser.add_argument("--sem-ordem", help="Expressão que identifica a mensagem de busca vazia")
    args = parser.parse_args()
    if args.planilha is None:
        args.planilha = input("Caminho da planilha .xlsx: ").strip().strip('"')
    if not args.planilha:
        parser.error("Informe o caminho da planilha.")
    if args.url is None:
        args.url = input("Endereço do SISBAJUD (https://...): ").strip()
        if not args.url.startswith("https://"):
            parser.error("Informe o endereço completo do SISBAJUD, começando com https://.")
    executar(args)
