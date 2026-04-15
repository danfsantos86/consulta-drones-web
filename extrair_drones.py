import re
import docx2txt
import os


ARQUIVO_DOCX = "Lista de Drones importados.docx"


def limpar_linha(texto):
    if texto is None:
        return ""
    texto = str(texto).replace("\r", " ").replace("\n", " ").strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def eh_frequencia(texto):
    t = texto.lower().replace(" ", "")
    return (
        "-" in texto
        or "mhz" in t
        or "a" in texto
    ) and any(ch.isdigit() for ch in texto)


def eh_potencia(texto):
    t = texto.strip().upper().replace(",", ".")
    if t == "ND":
        return True
    return bool(re.fullmatch(r"[0-9.]+", t))


def extrair_linhas_relevantes(texto):
    linhas = [limpar_linha(l) for l in texto.splitlines()]
    linhas = [l for l in linhas if l]

    ignorar_exatos = {
        "Lista de Drones Importados",
        "FABRICANTE",
        "MODELO",
        "NOME COMERCIAL",
        "FAIXA DE FREQUÊNCIA TX (MHz)",
        "POTÊNCIA MÁXIMA DE SAÍDA (W)",
    }

    resultado = []
    for linha in linhas:
        if linha in ignorar_exatos:
            continue
        if "Os modelos de Drones listados abaixo" in linha:
            continue
        if "Acesse:" in linha:
            continue
        if "gov.br/anatel" in linha.lower():
            continue
        resultado.append(linha)

    return resultado


def montar_registros(linhas):
    registros = []
    i = 0
    total = len(linhas)

    while i < total - 2:
        fabricante = linhas[i]
        modelo = linhas[i + 1]
        nome_comercial = linhas[i + 2]

        j = i + 3

        while j < total and eh_frequencia(linhas[j]):
            j += 1

        while j < total and eh_potencia(linhas[j]):
            j += 1

        registros.append({
            "FABRICANTE": fabricante,
            "MODELO": modelo,
            "NOME COMERCIAL": nome_comercial
        })

        i = j

    return registros


def carregar_drones():
    if not os.path.exists(ARQUIVO_DOCX):
        raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO_DOCX}")

    texto = docx2txt.process(ARQUIVO_DOCX)

    if not texto or not texto.strip():
        raise ValueError("Não foi possível extrair texto do arquivo DOCX.")

    linhas = extrair_linhas_relevantes(texto)
    registros = montar_registros(linhas)

    if not registros:
        raise ValueError("Nenhum registro foi identificado.")

    return registros


if __name__ == "__main__":
    drones = carregar_drones()
    print(f"Total de registros carregados: {len(drones)}")