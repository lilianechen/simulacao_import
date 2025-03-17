import streamlit as st
import pandas as pd
import shutil
from openpyxl import load_workbook
import os

# Caminho do arquivo original
ORIGINAL_FILE_PATH = "/mnt/data/001_AP_SIMULAÇÃ̃O_AMIGÃO_TESTE_container_papelariaBBR.xlsx"
SIMULATION_FILE_PATH = "/mnt/data/simulacao.xlsx"

def create_simulation_copy():
    """Cria uma cópia da planilha original para simulação."""
    shutil.copy(ORIGINAL_FILE_PATH, SIMULATION_FILE_PATH)

def update_excel(input_data):
    """Atualiza os valores na cópia da planilha e extrai os resultados."""
    wb = load_workbook(SIMULATION_FILE_PATH)
    
    # Atualizar células na aba "Dados Gerais"
    ws_dados_gerais = wb["Dados Gerais"]
    for cell, value in input_data.get("Dados Gerais", {}).items():
        if value is not None:
            ws_dados_gerais[cell] = value
    
    # Atualizar células na aba "Adições"
    ws_adicoes = wb["Adições"]
    for cell, value in input_data.get("Adições", {}).items():
        if value is not None:
            ws_adicoes[cell] = value
    
    # Salvar e permitir cálculos no Excel
    wb.save(SIMULATION_FILE_PATH)
    
    # Extrair resultados da aba "OUTPUT"
    ws_output = wb["OUTPUT"]  # Ajuste conforme necessário
    output_cells = ["Z5", "AA5", "AB5", "AC5", "AD5", "AE5", "AF5", "AG5"]  # Defina as células corretas
    output_data = {cell: ws_output[cell].value for cell in output_cells}
    
    return output_data

# Criar interface no Streamlit
st.title("Simulação de Importação")

# Criar cópia da planilha para cada execução
create_simulation_copy()

# Inputs do usuário
input_data = {"Dados Gerais": {}, "Adições": {}}

# Perguntar primeiro os dados da importação (aba "Dados Gerais")
st.header("Dados da Importação")
input_data["Dados Gerais"]["C21"] = st.number_input("Frete Internacional (C21)", value=0.0)
input_data["Dados Gerais"]["C28"] = st.number_input("Seguro (C28)", value=0.0)
input_data["Dados Gerais"]["C41"] = st.number_input("Custo Despacho (C41)", value=0.0)
input_data["Dados Gerais"]["C42"] = st.number_input("Custo Armazenagem (C42)", value=0.0)
input_data["Dados Gerais"]["C57"] = st.number_input("Outros Custos (C57)", value=0.0)

# Perguntar os dados dos produtos e suas informações (aba "Adições")
st.header("Produtos e Informações")
for i in range(5, 44):  # Vai até a linha 43
    input_data["Adições"][f"Z{i}"] = st.number_input(f"Produto {i-4} - Valor Unitário (Z{i})", value=0.0)
    input_data["Adições"][f"AA{i}"] = st.number_input(f"Produto {i-4} - Valor Total (AA{i})", value=0.0)
    input_data["Adições"][f"AB{i}"] = st.number_input(f"Produto {i-4} - ICMS (AB{i})", value=0.0)
    input_data["Adições"][f"AC{i}"] = st.number_input(f"Produto {i-4} - IPI (AC{i})", value=0.0)
    
# Remover linhas vazias da aba "Adições"
input_data["Adições"] = {cell: value for cell, value in input_data["Adições"].items() if value != 0.0}

if st.button("Executar Simulação"):
    result = update_excel(input_data)
    st.write("Resultados da Simulação:")
    for cell, value in result.items():
        st.write(f"{cell}: {value}")
