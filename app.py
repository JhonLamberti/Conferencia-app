import re
import io
import unicodedata
from difflib import get_close_matches
from typing import Dict, List, Tuple, Optional

import pandas as pd
import pdfplumber
import streamlit as st


EVENTOS = {
    "Hora Extra 75%": ["hora extra com 75%", "hora extra 75%"],
    "Hora Extra 100%": ["hora extra 100%"],
    "Noturno 27%": ["noturno 27%"],
    "Faltas Por Hora": ["faltas por hora"],
    "Desconto DSR": ["desconto dsr"],
    "Faltas em Dia": ["faltas em dia"],
    "Repouso Remunerado": ["repouso remunerado"],
}

COLUNAS = ["Código", "Colaborador", "Página"]
for evento in EVENTOS:
    if evento == "Repouso Remunerado":
        COLUNAS += [f"{evento} Valor"]
    else:
        COLUNAS += [f"{evento} Ref", f"{evento} Valor"]

COLUNAS_COMPARACAO = [
    "Colaborador PDF",
    "Colaborador Excel",
    "Código",
    "Página",
    "HE 75% PDF",
    "HE 75% Excel",
    "Diferença HE 75%",
    "Status HE 75%",
    "HE 100% PDF",
    "HE 100% Excel",
    "Diferença HE 100%",
    "Status HE 100%",
    "Noturno PDF",
    "Noturno Excel",
    "Diferença Noturno",
    "Status Noturno",
    "Status Geral",
]


# -----------------------------
# Utilidades gerais
# -----------------------------

def normalizar_texto(texto: str) -> str:
    texto = unicodedata.normalize("NFD", str(texto or ""))
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    texto = re.sub(r"\s+", " ", texto)
    return texto.lower().strip()


def normalizar_nome(nome: str) -> str:
    nome = normalizar_texto(nome)
    nome = re.sub(r"[^a-z0-9 ]", "", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome


def limpar_nome(nome: str) -> str:
    nome = re.sub(r"\s+", " ", nome or "").strip()
    nome = re.split(r"\s+\d{1,3}[.,]\d{2,}", nome)[0].strip()
    return nome


def tempo_para_minutos(valor) -> int:
    """Converte formatos como 001:00, 33:49, 00:00 Not.: 00:00 para minutos."""
    if valor is None:
        return 0
    texto = str(valor).strip()
    if not texto:
        return 0

    # Pega sempre o primeiro HH:MM da célula. Ex.: "06:00 Not.: 00:00" => 06:00
    m = re.search(r"(\d{1,4}):(\d{2})", texto)
    if not m:
        return 0
    return int(m.group(1)) * 60 + int(m.group(2))


def minutos_para_tempo(minutos: int) -> str:
    sinal = "-" if minutos < 0 else ""
    minutos = abs(int(minutos))
    return f"{sinal}{minutos // 60:02d}:{minutos % 60:02d}"


def status_por_diferenca(diff_minutos: int) -> str:
    return "OK" if diff_minutos == 0 else "DIVERGENTE"


# -----------------------------
# Extração do PDF
# -----------------------------

def extrair_colaboradores_do_texto(texto: str, pagina: int) -> List[Dict]:
    """
    Divide o texto da página em blocos por colaborador.
    O padrão do PDF é uma linha começando com código de 6 dígitos + nome.
    Ex.: 000321 BENTO DA SILVA FERREIRA 5.685,12 320 0320
    """
    linhas = [l.strip() for l in (texto or "").splitlines() if l.strip()]
    inicios: List[Tuple[int, str, str]] = []

    padrao_colaborador = re.compile(
        r"^(\d{6})\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ ]{5,}?)(?=\s+\d{1,3}(?:\.\d{3})*,\d{2}|\s+\d{3,4}\s|\s+Admissão|$)"
    )

    for idx, linha in enumerate(linhas):
        m = padrao_colaborador.search(linha)
        if m:
            codigo = m.group(1)
            nome = limpar_nome(m.group(2))
            if nome and "EMPRESA" not in nome and "FOLHA" not in nome:
                inicios.append((idx, codigo, nome))

    colaboradores = []
    for pos, (idx, codigo, nome) in enumerate(inicios):
        fim = inicios[pos + 1][0] if pos + 1 < len(inicios) else len(linhas)
        bloco = "\n".join(linhas[idx:fim])
        colaboradores.append({
            "Código": codigo,
            "Colaborador": nome,
            "Página": pagina,
            "bloco": bloco,
        })
    return colaboradores


def extrair_evento_do_bloco(bloco: str, evento: str, aliases: List[str]) -> Tuple[str, str]:
    """
    Extrai referência e valor do evento dentro do bloco do colaborador.

    Importante: dependendo do modo de extração do PDF, a linha pode sair como:
      404 Hora Extra com 75% 001:00 43,76
    ou como:
      404 Hora Extra com 75% 43,76 001:00

    Por isso a função não depende da posição. Ela procura separadamente:
    - referência/hora no padrão HH:MM, 001:00, 033:49 etc.
    - valor em reais no padrão 43,76, 1.691,37 etc.
    """
    for linha in bloco.splitlines():
        linha_norm = normalizar_texto(linha)
        if any(alias in linha_norm for alias in aliases):
            horas = re.findall(r"\b\d{1,4}:\d{2}\b", linha)
            valores = re.findall(r"\b\d{1,3}(?:\.\d{3})*,\d{2}\b", linha)

            if evento == "Repouso Remunerado":
                return "", valores[-1] if valores else ""

            referencia = horas[-1] if horas else ""
            valor = valores[-1] if valores else ""
            return referencia, valor

    return "", ""


def processar_pdf(arquivo_pdf) -> pd.DataFrame:
    registros = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        for numero_pagina, page in enumerate(pdf.pages, start=1):
            texto = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            colaboradores = extrair_colaboradores_do_texto(texto, numero_pagina)
            for colab in colaboradores:
                linha = {col: "" for col in COLUNAS}
                linha["Código"] = colab["Código"]
                linha["Colaborador"] = colab["Colaborador"]
                linha["Página"] = colab["Página"]
                bloco = colab["bloco"]
                for evento, aliases in EVENTOS.items():
                    ref, valor = extrair_evento_do_bloco(bloco, evento, aliases)
                    if evento == "Repouso Remunerado":
                        linha[f"{evento} Valor"] = valor
                    else:
                        linha[f"{evento} Ref"] = ref
                        linha[f"{evento} Valor"] = valor
                registros.append(linha)
    return pd.DataFrame(registros, columns=COLUNAS)


# -----------------------------
# Extração da planilha de ponto
# -----------------------------

def processar_planilha_ponto(arquivo_excel) -> pd.DataFrame:
    """
    Lê a planilha de ponto em blocos:
    - Linha com A = Colaborador e B = Nome
    - Linha TOTAIS do mesmo bloco
    - Colunas J, K e L: HE 75%, HE 100%, Adicional Noturno
    """
    df_raw = pd.read_excel(arquivo_excel, header=None, dtype=str, engine="openpyxl").fillna("")
    registros = []

    colaborador_atual = None
    for idx, row in df_raw.iterrows():
        col_a = str(row.iloc[0]).strip()
        col_b = str(row.iloc[1]).strip() if len(row) > 1 else ""

        if normalizar_texto(col_a) == "colaborador" and col_b:
            colaborador_atual = col_b
            continue

        if colaborador_atual and normalizar_texto(col_a) == "totais":
            he75 = row.iloc[9] if len(row) > 9 else ""
            he100 = row.iloc[10] if len(row) > 10 else ""
            noturno = row.iloc[11] if len(row) > 11 else ""

            registros.append({
                "Colaborador Excel": colaborador_atual,
                "Chave Nome": normalizar_nome(colaborador_atual),
                "HE 75% Excel": minutos_para_tempo(tempo_para_minutos(he75)),
                "HE 100% Excel": minutos_para_tempo(tempo_para_minutos(he100)),
                "Noturno Excel": minutos_para_tempo(tempo_para_minutos(noturno)),
                "HE 75% Excel Min": tempo_para_minutos(he75),
                "HE 100% Excel Min": tempo_para_minutos(he100),
                "Noturno Excel Min": tempo_para_minutos(noturno),
            })
            colaborador_atual = None

    return pd.DataFrame(registros)


def encontrar_colaborador_excel(nome_pdf: str, mapa_excel: Dict[str, dict]) -> Optional[dict]:
    chave = normalizar_nome(nome_pdf)
    if chave in mapa_excel:
        return mapa_excel[chave]

    # Plano B: correspondência aproximada para casos com acento, cedilha ou pequenas variações.
    chaves = list(mapa_excel.keys())
    similares = get_close_matches(chave, chaves, n=1, cutoff=0.92)
    if similares:
        return mapa_excel[similares[0]]
    return None


def comparar_pdf_com_excel(df_pdf: pd.DataFrame, df_excel: pd.DataFrame) -> pd.DataFrame:
    if df_excel.empty:
        return pd.DataFrame(columns=COLUNAS_COMPARACAO)

    mapa_excel = {row["Chave Nome"]: row for _, row in df_excel.iterrows()}
    registros = []

    for _, row_pdf in df_pdf.iterrows():
        nome_pdf = row_pdf.get("Colaborador", "")
        match = encontrar_colaborador_excel(nome_pdf, mapa_excel)

        pdf_he75_min = tempo_para_minutos(row_pdf.get("Hora Extra 75% Ref", ""))
        pdf_he100_min = tempo_para_minutos(row_pdf.get("Hora Extra 100% Ref", ""))
        pdf_noturno_min = tempo_para_minutos(row_pdf.get("Noturno 27% Ref", ""))

        if match is None:
            registros.append({
                "Colaborador PDF": nome_pdf,
                "Colaborador Excel": "NÃO ENCONTRADO",
                "Código": row_pdf.get("Código", ""),
                "Página": row_pdf.get("Página", ""),
                "HE 75% PDF": minutos_para_tempo(pdf_he75_min),
                "HE 75% Excel": "",
                "Diferença HE 75%": "",
                "Status HE 75%": "NÃO ENCONTRADO",
                "HE 100% PDF": minutos_para_tempo(pdf_he100_min),
                "HE 100% Excel": "",
                "Diferença HE 100%": "",
                "Status HE 100%": "NÃO ENCONTRADO",
                "Noturno PDF": minutos_para_tempo(pdf_noturno_min),
                "Noturno Excel": "",
                "Diferença Noturno": "",
                "Status Noturno": "NÃO ENCONTRADO",
                "Status Geral": "COLABORADOR NÃO ENCONTRADO NO EXCEL",
            })
            continue

        excel_he75_min = int(match.get("HE 75% Excel Min", 0))
        excel_he100_min = int(match.get("HE 100% Excel Min", 0))
        excel_noturno_min = int(match.get("Noturno Excel Min", 0))

        diff_he75 = pdf_he75_min - excel_he75_min
        diff_he100 = pdf_he100_min - excel_he100_min
        diff_noturno = pdf_noturno_min - excel_noturno_min

        status_he75 = status_por_diferenca(diff_he75)
        status_he100 = status_por_diferenca(diff_he100)
        status_noturno = status_por_diferenca(diff_noturno)
        status_geral = "OK" if all(s == "OK" for s in [status_he75, status_he100, status_noturno]) else "DIVERGENTE"

        registros.append({
            "Colaborador PDF": nome_pdf,
            "Colaborador Excel": match.get("Colaborador Excel", ""),
            "Código": row_pdf.get("Código", ""),
            "Página": row_pdf.get("Página", ""),
            "HE 75% PDF": minutos_para_tempo(pdf_he75_min),
            "HE 75% Excel": minutos_para_tempo(excel_he75_min),
            "Diferença HE 75%": minutos_para_tempo(diff_he75),
            "Status HE 75%": status_he75,
            "HE 100% PDF": minutos_para_tempo(pdf_he100_min),
            "HE 100% Excel": minutos_para_tempo(excel_he100_min),
            "Diferença HE 100%": minutos_para_tempo(diff_he100),
            "Status HE 100%": status_he100,
            "Noturno PDF": minutos_para_tempo(pdf_noturno_min),
            "Noturno Excel": minutos_para_tempo(excel_noturno_min),
            "Diferença Noturno": minutos_para_tempo(diff_noturno),
            "Status Noturno": status_noturno,
            "Status Geral": status_geral,
        })

    return pd.DataFrame(registros, columns=COLUNAS_COMPARACAO)


# -----------------------------
# Exportação
# -----------------------------

def ajustar_larguras(ws):
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 40)


def gerar_excel(df_pdf: pd.DataFrame, df_excel: Optional[pd.DataFrame] = None, df_comparacao: Optional[pd.DataFrame] = None) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_pdf.to_excel(writer, index=False, sheet_name="Eventos PDF")
        ajustar_larguras(writer.book["Eventos PDF"])
        writer.book["Eventos PDF"].freeze_panes = "A2"

        if df_excel is not None and not df_excel.empty:
            df_excel_export = df_excel.drop(columns=[c for c in df_excel.columns if c.endswith(" Min") or c == "Chave Nome"], errors="ignore")
            df_excel_export.to_excel(writer, index=False, sheet_name="Totais Ponto")
            ajustar_larguras(writer.book["Totais Ponto"])
            writer.book["Totais Ponto"].freeze_panes = "A2"

        if df_comparacao is not None and not df_comparacao.empty:
            df_comparacao.to_excel(writer, index=False, sheet_name="Comparacao")
            ws = writer.book["Comparacao"]
            ajustar_larguras(ws)
            ws.freeze_panes = "A2"
    return output.getvalue()


# -----------------------------
# Interface Streamlit
# -----------------------------

st.set_page_config(page_title="Conferência de Folha PDF x Ponto", layout="wide")
st.title("Conferência de eventos da folha em PDF x planilha de ponto")
st.write("Envie o PDF da folha e, opcionalmente, a planilha de ponto para comparar HE 75%, HE 100% e adicional noturno.")

arquivo_pdf = st.file_uploader("1. Selecione o PDF da folha", type=["pdf"])
arquivo_excel = st.file_uploader("2. Selecione a planilha de ponto padrão", type=["xlsx"], help="A planilha deve ter blocos por colaborador e linha TOTAIS com colunas J, K e L.")

if arquivo_pdf:
    if st.button("Processar conferência", type="primary"):
        try:
            with st.spinner("Lendo PDF e identificando colaboradores/eventos..."):
                df_pdf = processar_pdf(arquivo_pdf)

            if df_pdf.empty:
                st.warning("Nenhum colaborador foi identificado no PDF. Verifique se o PDF possui texto extraível.")
            else:
                st.success(f"PDF processado. {len(df_pdf)} colaboradores identificados.")

                colunas_eventos = [c for c in df_pdf.columns if c not in ["Código", "Colaborador", "Página"]]
                df_pdf_com_evento = df_pdf[df_pdf[colunas_eventos].apply(lambda row: any(str(v).strip() for v in row), axis=1)]

                st.subheader("Eventos encontrados no PDF")
                st.dataframe(df_pdf_com_evento if not df_pdf_com_evento.empty else df_pdf, use_container_width=True)

                df_excel = None
                df_comparacao = None

                if arquivo_excel:
                    with st.spinner("Lendo planilha de ponto e comparando com o PDF..."):
                        df_excel = processar_planilha_ponto(arquivo_excel)
                        df_comparacao = comparar_pdf_com_excel(df_pdf, df_excel)

                    st.subheader("Totais identificados na planilha de ponto")
                    if df_excel.empty:
                        st.warning("Nenhum total foi identificado na planilha. Verifique se existe linha 'Colaborador' e linha 'TOTAIS'.")
                    else:
                        df_excel_view = df_excel.drop(columns=[c for c in df_excel.columns if c.endswith(" Min") or c == "Chave Nome"], errors="ignore")
                        st.dataframe(df_excel_view, use_container_width=True)

                    st.subheader("Comparação PDF x Planilha")
                    if df_comparacao is not None and not df_comparacao.empty:
                        apenas_divergencias = df_comparacao[df_comparacao["Status Geral"] != "OK"]
                        st.metric("Divergências encontradas", len(apenas_divergencias))
                        st.dataframe(apenas_divergencias if not apenas_divergencias.empty else df_comparacao, use_container_width=True)
                    else:
                        st.info("Nenhuma comparação gerada.")

                excel_bytes = gerar_excel(df_pdf, df_excel, df_comparacao)
                st.download_button(
                    label="Baixar Excel da conferência",
                    data=excel_bytes,
                    file_name="conferencia_pdf_x_ponto.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error(f"Erro ao processar os arquivos: {e}")

st.divider()
st.caption("Eventos do PDF: HE 75%, HE 100%, Noturno 27%, Faltas por Hora, Desconto DSR, Faltas em Dia e Repouso Remunerado. Comparação com Excel: HE 75%, HE 100% e adicional noturno.")
