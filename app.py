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
st.write("Faça o upload dos PDFs para extrair os dados das bases **1223**, **1247** e **1262**.")

# O parâmetro accept_multiple_files=True permite selecionar vários ficheiros
uploaded_files = st.file_uploader("Arraste e solte os arquivos PDF da CPFL aqui", type=["pdf"], accept_multiple_files=True)

if uploaded_files: # Verifica se a lista de ficheiros não está vazia
    with st.spinner("A processar os ficheiros PDF..."):
        
        city_pattern = re.compile(r'^#{1,4}\s*(.+)')
        ul_pattern = re.compile(r'Unidade de Leitura \(UL\):\s*(.+)', re.IGNORECASE)
        data_row_pattern = re.compile(r'^(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)$')
        
        target_bases = ['1223', '1247', '1262']
        all_rows = [] # Lista para acumular os dados de todos os PDFs
        
        # Itera sobre cada ficheiro carregado
        for uploaded_file in uploaded_files:
            reader = pypdf.PdfReader(uploaded_file)
            
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
                            all_rows.append({
                                'Arquivo Origem': uploaded_file.name, # Adiciona o nome do ficheiro para rastreio
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
        
        if len(all_rows) > 0:
            df = pd.DataFrame(all_rows)
            st.success(f"Concluído! Foram encontrados {len(df)} registos em {len(uploaded_files)} ficheiro(s).")
            st.dataframe(df.head(20))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Bases Operacionais')
            
            st.download_button(
                label="📥 Descarregar Planilha Excel Unificada",
                data=output.getvalue(),
                file_name="Relatorio_Bases_CPFL_Unificado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum registo foi encontrado para as bases selecionadas nos ficheiros fornecidos.")
