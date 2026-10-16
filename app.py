import subprocess
import sys

for package in ["pypdf", "openpyxl"]:
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Processador CPFL", layout="wide")

st.title("📊 Processador de Manuseio CPFL")
st.write("Faça o upload do PDF para extrair os dados das bases **1223**, **1247** e **1262**.")

uploaded_file = st.file_uploader("Arraste e solte o arquivo PDF da CPFL aqui", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("A processar o ficheiro PDF..."):
        reader = pypdf.PdfReader(uploaded_file)
        
        city_pattern = re.compile(r'^#{1,4}\s*(.+)')
        ul_pattern = re.compile(r'Unidade de Leitura \(UL\):\s*(.+)', re.IGNORECASE)
        data_row_pattern = re.compile(r'^(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)$')
        
        target_bases = ['1223', '1247', '1262']
        rows = []
        curr_base = None
        curr_city = None
        curr_ul = None
        
        for page_idx, page in enumerate(reader.pages):
            text = page.extract_text()
            if not text:
                continue
                
            lines = text.split('\n')
            for line in lines:
                line_s = line.strip()
                if not line_s:
                    continue
                
                if re.match(r'^\d{4}\s*-', line_s):
                    curr_base = line_s
                    continue
                    
                m_city = city_pattern.match(line_s)
                if m_city:
                    curr_city = m_city.group(1).strip()
                    continue
                    
                m_ul = ul_pattern.search(line_s)
                if m_ul:
                    curr_ul = m_ul.group(1).strip()
                    continue
                    
                m_data = data_row_pattern.match(line_s)
                if m_data:
                    if curr_base and any(base_code in curr_base for base_code in target_bases):
                        cod_entr, nro_pacote, nu_inicial, nu_final, qtde_contas = m_data.groups()
                        rows.append({
                            'Base Operacional': curr_base,
                            'Cidade': curr_city,
                            'Cod Entr': cod_entr,
                            'Nro Pacote': nro_pacote,
                            'Nu Impr Inicial': int(nu_inicial),
                            'Nu Impr Final': int(nu_final),
                            'Qtde de Contas': int(qtde_contas),
                            'Unidade de Leitura (UL)': curr_ul,
                            'Página PDF': page_idx + 1
                        })
        
        if len(rows) > 0:
            df = pd.DataFrame(rows)
            st.success(f"Concluído! Foram encontrados {len(df)} registos.")
            st.dataframe(df.head(20))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Bases Operacionais')
            
            st.download_button(
                label="📥 Descarregar Planilha Excel",
                data=output.getvalue(),
                file_name="Relatorio_Bases_CPFL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum registo foi encontrado para as bases selecionadas.")
