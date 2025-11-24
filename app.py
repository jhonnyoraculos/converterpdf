import re
import io
import pdfplumber
import pandas as pd
import streamlit as st
from datetime import datetime

# Colunas fixas solicitadas
COLUMNS = [
    "rota",
    "data_emissao",
    "data_previsao",
    "motorista",
    "veiculo",
    "carga",
    "numero_nota",
    "codigo_cliente",
    "nome_cliente",
    "pedido",
    "cidade",
    "peso_pedido",
    "endereco",
    "total_nota",
    "forma_recebimento",
    "valor_recebimento",
]


def parse_br_number(value):
    """Converte números no padrão brasileiro (5.458,96) para float."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    normalized = text.replace(".", "").replace(",", ".")
    try:
        return float(normalized)
    except ValueError:
        return None


def extract_header_fields(text: str) -> dict:
    """Extrai campos gerais do romaneio que valem para todas as notas."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    rota = ""
    for line in lines:
        # Procura primeira linha com número + nome de rota (ex.: "600 PEDRO LEOPOLDO")
        if re.match(r"^\d{2,4}\s+[A-Za-zÀ-ÿ0-9 .,\-()]+$", line):
            rota = line
            break

    def find(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    data_emissao = find(r"Emiss[aã]o:\s*([\d/]{8,10})")
    data_previsao = find(r"Previs[aã]o:\s*([\d/]{8,10})")
    motorista = find(r"Motorista:\s*([^\n]+)")
    veiculo = find(r"Ve[ií]culo:\s*([^\n]+)")
    carga = find(r"Carga[:\s]+([0-9\.,]+)")

    return {
        "rota": rota,
        "data_emissao": data_emissao,
        "data_previsao": data_previsao,
        "motorista": motorista,
        "veiculo": veiculo,
        "carga": carga,
    }


def split_notes_blocks(text: str) -> list:
    """Divide o texto bruto em blocos de notas identificando possíveis números de nota."""
    # Primeiro tenta o padrão mais comum (ex.: 1552-24995 ou 748-12263)
    note_pattern = re.compile(r"(?:^|\n)\s*(\d{3,5}-\d{3,})")
    positions = [m.start(1) for m in note_pattern.finditer(text)]

    # Se não achar nada, tenta usar "Pedido:" como delimitador de blocos
    if not positions:
        pedido_pattern = re.compile(r"(?:^|\n)\s*Pedido:\s*\d+")
        positions = [m.start() for m in pedido_pattern.finditer(text)]

    if not positions:
        return []

    blocks = []
    for idx, start in enumerate(positions):
        end = positions[idx + 1] if idx + 1 < len(positions) else len(text)
        blocks.append(text[start:end].strip())
    return blocks


def parse_block(block: str, header: dict) -> dict:
    """Extrai campos de um bloco individual de nota."""

    def find(pattern):
        match = re.search(pattern, block, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    numero_nota = find(r"^\s*(\d{3,5}-\d+)")

    nome_line = find(r"Nome:\s*([^\n]+)")
    codigo_cliente, nome_cliente = "", ""
    if nome_line:
        code_match = re.match(r"(\d+)\s*-\s*(.+)", nome_line)
        if code_match:
            codigo_cliente = code_match.group(1).strip()
            nome_cliente = code_match.group(2).strip()
        else:
            nome_cliente = nome_line

    pedido = find(r"Pedido:\s*([^\n]+)")
    cidade = find(r"Cidade:\s*([^\n]+)")
    peso_pedido = parse_br_number(find(r"Peso\s*Pedido:\s*([0-9\.,]+)"))
    endereco = find(r"Endere[cç]o:\s*([^\n]+)")
    total_nota = parse_br_number(find(r"Total\s+da\s+Nota:\s*R?\$?\s*([0-9\.,]+)"))

    forma_recebimento, valor_recebimento = "", None
    dup_match = re.search(
        r"Duplicata\s+a\s+Receber\s*(.*?)\s*Valor:\s*R?\$?\s*([0-9\.,]+)",
        block,
        re.IGNORECASE | re.DOTALL,
    )
    if dup_match:
        forma_recebimento = dup_match.group(1).strip()
        valor_recebimento = parse_br_number(dup_match.group(2))

    record = {
        **header,
        "numero_nota": numero_nota,
        "codigo_cliente": codigo_cliente,
        "nome_cliente": nome_cliente,
        "pedido": pedido,
        "cidade": cidade,
        "peso_pedido": peso_pedido,
        "endereco": endereco,
        "total_nota": total_nota,
        "forma_recebimento": forma_recebimento,
        "valor_recebimento": valor_recebimento,
    }

    # Garante chaves com valores vazios quando não encontrados
    for key in COLUMNS:
        if key not in record:
            record[key] = None

    return record


def parse_text_to_records(text: str) -> list:
    """Recebe texto bruto do PDF e devolve lista de dicionários das notas."""
    header = extract_header_fields(text)
    blocks = split_notes_blocks(text)
    records = []

    for block in blocks:
        records.append(parse_block(block, header))

    return records


def parse_pdf(file) -> pd.DataFrame:
    """Abre um PDF, concatena o texto das páginas e retorna DataFrame das notas."""
    full_text_parts = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text_parts.append(page_text)
    full_text = "\n".join(full_text_parts)

    records = parse_text_to_records(full_text)
    df = pd.DataFrame(records, columns=COLUMNS)
    return df


def format_date_br(date_str: str) -> str:
    """Normaliza datas dd/mm/yyyy ou dd/mm/yy (opcional)."""
    if not date_str:
        return ""
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(date_str, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return date_str.strip()


def main():
    st.set_page_config(page_title="Gerador de Planilha - Romaneio JR Ferragens", layout="wide")
    st.title("GERADOR DE PLANILHA - ROMANEIO JR FERRAGENS")
    st.write("Faça upload de um ou mais PDFs do romaneio para gerar a planilha consolidada.")

    uploaded_files = st.file_uploader(
        "Selecione os arquivos PDF",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Aguardando arquivos PDF para processar.")
        return

    dataframes = []
    for uploaded_file in uploaded_files:
        try:
            df_pdf = parse_pdf(uploaded_file)
            if not df_pdf.empty:
                # Normaliza datas para um formato único
                df_pdf["data_emissao"] = df_pdf["data_emissao"].apply(format_date_br)
                df_pdf["data_previsao"] = df_pdf["data_previsao"].apply(format_date_br)
                dataframes.append(df_pdf)
            else:
                st.warning(f"Nenhuma nota encontrada em {uploaded_file.name}.")
        except Exception as exc:
            st.error(f"Erro ao processar {uploaded_file.name}: {exc}")

    if not dataframes:
        st.warning("Nenhuma nota foi identificada nos PDFs enviados.")
        return

    df_all = pd.concat(dataframes, ignore_index=True)

    st.subheader("Notas encontradas")
    st.dataframe(df_all)

    total_nota_sum = pd.to_numeric(df_all["total_nota"], errors="coerce").fillna(0).sum()
    peso_total_sum = pd.to_numeric(df_all["peso_pedido"], errors="coerce").fillna(0).sum()

    col1, col2 = st.columns(2)
    col1.metric("Soma das Notas (R$)", f"{total_nota_sum:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    col2.metric("Peso Total", f"{peso_total_sum:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Gera Excel em memória para download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_all.to_excel(writer, index=False, sheet_name="Romaneio")
    buffer.seek(0)

    rota_name = df_all["rota"].dropna().astype(str).str.strip()
    rota_label = rota_name.iloc[0] if not rota_name.empty and rota_name.iloc[0] else "romaneio"
    file_safe = re.sub(r"[^A-Za-z0-9\-]+", "_", rota_label).strip("_") or "romaneio"
    download_name = f"romaneio_rota_{file_safe}.xlsx"

    st.download_button(
        label="Baixar Excel",
        data=buffer,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
