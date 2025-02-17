import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
from io import BytesIO

# Função para processar os dados com base no tipo de DataLogger
def processar_dados(tipo_datalogger, arquivo):
    try:
        # Lendo o arquivo enviado
        extensao = os.path.splitext(arquivo.name)[-1]
        if extensao == ".csv":
            df = pd.read_csv(arquivo, sep=";", skiprows=5)
        elif extensao in [".xls", ".xlsx"]:
            df = pd.read_excel(arquivo)

        # Processamento baseado no tipo de DataLogger
        if tipo_datalogger == "Lamon":
            if 'Hora' in df.columns:
                df['Hora'], df['Minutos'], df['Segundos'] = zip(
                    *df['Hora'].str.split(':').apply(lambda x: (x[0].zfill(2), x[1].zfill(2), x[2].zfill(2)))
                )
            if 'Pressão(mca)' in df.columns:
                df['Pressão(mca)'] = df['Pressão(mca)'].astype(float).round(2)
                df['Pressão_mca'] = df['Pressão(mca)'].apply(lambda x: f"{x:.2f}")
                df.drop(columns=['Pressão(mca)'], inplace=True)
        elif tipo_datalogger == "Sanesoluti":
            if 'Hora' in df.columns:
                df['Hora'], df['Minutos'], df['Segundos'] = zip(
                    *df['Hora'].str.split(':').apply(lambda x: (x[0].zfill(2), x[1].zfill(2), x[2].zfill(2)))
                )
            if 'Pressão' in df.columns and 'Volume Total' in df.columns:
                df['Pressão'] = df['Pressão'].astype(float).round(2)
                df['Volume Total'] = df['Volume Total'].astype(float).round(2)

        return df

    except Exception as e:
        st.error(f"Erro ao processar os dados: {e}")
        return None

# Função para converter DataFrame em Excel com tabela formatada
def converter_para_excel_com_tabela(df):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active

    # Adicionar cabeçalhos na planilha
    for col_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=1, column=col_idx, value=col_name)

    # Adicionar dados do DataFrame na planilha
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # Criar tabela formatada
    ref = f"A1:{chr(64 + len(df.columns))}{len(df) + 1}"  # Define o intervalo da tabela (ex: A1:C4)
    tabela = Table(displayName="Tabela1", ref=ref)
    estilo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=True,
    )
    tabela.tableStyleInfo = estilo
    ws.add_table(tabela)

    # Salvar o arquivo na memória
    wb.save(output)
    return output.getvalue()

# Interface do Streamlit
st.title("Processador de Arquivos DataLogger")

# Upload do arquivo
arquivo = st.file_uploader("Faça upload do arquivo (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"])

# Seleção do tipo de DataLogger
tipos_datalogger = ["Lamon", "Sanesoluti", "Vectora"]
tipo_datalogger = st.selectbox("Selecione o tipo de DataLogger", tipos_datalogger)

# Botão para processar os dados
if st.button("Processar"):
    if arquivo is not None and tipo_datalogger:
        # Processar os dados
        resultado_df = processar_dados(tipo_datalogger, arquivo)
        
        if resultado_df is not None:
            # Exibir os dados processados
            st.success("Dados processados com sucesso!")
            st.dataframe(resultado_df)

            # Converter o DataFrame para Excel com tabela formatada e permitir download
            excel_data = converter_para_excel_com_tabela(resultado_df)
            st.download_button(
                label="Baixar arquivo processado como Tabela",
                data=excel_data,
                file_name="arquivo_processado_tabela.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Por favor, faça upload de um arquivo e selecione um tipo de DataLogger.")
