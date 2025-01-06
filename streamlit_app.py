import os
from groq import Groq
import streamlit as st
import pandas as pd
import openpyxl
import unidecode
import json
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px

# Adicionar após as importações
MODELOS_DISPONIVEIS = {
    # Modelos Estáveis
    "Gemma 2 9B": "gemma2-9b-it",
    "LLama 3.3 70B Versatile": "llama-3.3-70b-versatile",
    "LLama 3.1 8B Instant": "llama-3.1-8b-instant",
    "LLama 3 70B": "llama3-70b-8192",
    "LLama 3 8B": "llama3-8b-8192",
    "Mixtral 8x7B": "mixtral-8x7b-32768",
    
    # Modelos Preview
    "LLama 3.3 70B SpecDec (Preview)": "llama-3.3-70b-specdec",
    "LLama 3.2 1B (Preview)": "llama-3.2-1b-preview",
    "LLama 3.2 3B (Preview)": "llama-3.2-3b-preview",
    "LLama 3.2 11B Vision (Preview)": "llama-3.2-11b-vision-preview",
    "LLama 3.2 90B Vision (Preview)": "llama-3.2-90b-vision-preview"
}

# Configuração da API Groq
GROQ_API_KEY = "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"
client = Groq(api_key=GROQ_API_KEY)

# Configurações iniciais do Streamlit
st.title("Gerador de Planilhas")

# No início do arquivo, após as importações
if 'use_yes_no' not in st.session_state:
    st.session_state.use_yes_no = False

if 'df' not in st.session_state:
    st.session_state.df = None
if 'calculos_aplicados' not in st.session_state:
    st.session_state.calculos_aplicados = False

# Checkbox com estado persistente
use_yes_no = st.checkbox('Usar formato Sim/Não para listas', 
                        value=st.session_state.use_yes_no,
                        key='use_yes_no')

# Adicionar seletor de modelo
st.sidebar.title("Configurações do Modelo")

# Agrupar modelos por categoria
modelos_estaveis = {k: v for k, v in MODELOS_DISPONIVEIS.items() if "Preview" not in k}
modelos_preview = {k: v for k, v in MODELOS_DISPONIVEIS.items() if "Preview" in k}

categoria_modelo = st.sidebar.radio(
    "Categoria do Modelo:",
    ["Modelos Estáveis", "Modelos Preview"]
)

modelos_disponiveis = modelos_estaveis if categoria_modelo == "Modelos Estáveis" else modelos_preview

modelo_selecionado = st.sidebar.selectbox(
    "Selecione o modelo:",
    list(modelos_disponiveis.keys()),
    index=0
)

# Definir funções de cálculo disponíveis
CALCULOS = {
    'Média': lambda x: x.mean(),
    'Soma': lambda x: x.sum(),
    'Máximo': lambda x: x.max(),
    'Mínimo': lambda x: x.min(),
    'Desvio Padrão': lambda x: x.std(),
    'Mediana': lambda x: x.median(),
    'Percentual': lambda x: (x / x.sum()) * 100,
    'Moda': lambda x: x.mode().iloc[0] if not x.mode().empty else None
}

def aplicar_calculos(df, coluna, calculos_selecionados):
    """Aplica os cálculos selecionados na coluna"""
    for calculo in calculos_selecionados:
        if calculo in CALCULOS:
            nova_coluna = f"{coluna}_{calculo}"
            df[nova_coluna] = CALCULOS[calculo](df[coluna])
    return df

def process_prompt_to_data(prompt, use_yes_no):
    try:
        formatted_prompt = f"""
        Gere apenas dados JSON válidos sem texto adicional.
        Formato: {{"usuarios": [{{"nome": "...", "lista": [...]}}]}}
        Não inclua explicações, apenas o JSON.
        
        Descrição: {prompt}
        """

        response = client.chat.completions.create(
            model=MODELOS_DISPONIVEIS[modelo_selecionado],
            messages=[{"role": "user", "content": formatted_prompt}],
            temperature=0.7,
            max_tokens=2000
        )
        
        content = response.choices[0].message.content
        cleaned_data = clean_json_response(content)
        data = json.loads(cleaned_data)
        
        # Primeiro aplicar conversão Sim/Não se necessário
        if use_yes_no:
            data = convert_to_yes_no(data)
            
        # Depois normalizar os dados
        data = normalize_data(data)
        
        return data
        
    except Exception as e:
        st.error(f"Erro: {str(e)}")
        return None

# Função para limpar a resposta JSON
def clean_json_response(response):
    try:
        # Encontrar o JSON mais externo
        start = response.find('{')
        count = 0
        end = -1
        
        if start == -1:
            return '{}'
            
        for i in range(start, len(response)):
            if response[i] == '{':
                count += 1
            elif response[i] == '}':
                count -= 1
                if count == 0:
                    end = i + 1
                    break
                    
        if end == -1:
            return '{}'
            
        json_str = response[start:end]
        
        # Validar JSON
        json.loads(json_str)
        return json_str
        
    except Exception as e:
        st.error(f"Erro na limpeza do JSON: {e}")
        return '{}'

# Função para validar a estrutura do JSON
def validate_json_structure(data):
    try:
        # Validar estrutura esperada
        if not isinstance(data, dict):
            return False
            
        if not any(key in data for key in ['usuarios', 'tabela', 'data']):
            return False
            
        return True
        
    except Exception:
        return False

# Função para normalizar os dados
def normalize_data(data):
    """Normaliza os dados e formata cabeçalhos"""
    try:
        if isinstance(data, dict):
            for key in data.keys():
                if isinstance(data[key], list):
                    normalized_items = []
                    for item in data[key]:
                        normalized_item = {}
                        for k, v in item.items():
                            # Formatar chave substituindo underscores por espaços
                            col_name = unidecode.unidecode(k).replace('_', ' ').title()
                            col_name = unidecode.unidecode(k).replace('-', ' ').title()
                            
                            if isinstance(v, list):
                                if use_yes_no:
                                    # Modo Sim/Não para listas
                                    for val in set(v):
                                        col_name = unidecode.unidecode(val).replace('_', ' ').title()
                                        normalized_item[col_name] = 'Sim' if val in v else 'Não'
                                else:
                                    # Modo lista
                                    normalized_item[col_name] = ', '.join(map(str, v))
                            else:
                                normalized_item[col_name] = v
                        normalized_items.append(normalized_item)
                    data[key] = normalized_items
        return data
        
    except Exception as e:
        st.error(f"Erro ao normalizar dados: {e}")
        return data

def convert_to_yes_no(data):
    """Converte listas em formato Sim/Não"""
    if not isinstance(data, dict) or 'usuarios' not in data:
        return data
        
    modified_data = {'usuarios': []}
    
    for item in data['usuarios']:
        new_item = {}
        for k, v in item.items():
            if isinstance(v, list):
                # Criar colunas Sim/Não para cada valor único na lista
                unique_values = set(v)
                for val in unique_values:
                    new_key = unidecode.unidecode(str(val)).replace('_', ' ').title()
                    new_item[new_key] = 'Sim' if val in v else 'Não'
            else:
                new_item[k] = v
        modified_data['usuarios'].append(new_item)
    
    return modified_data

# Função para salvar os dados em Excel
def save_data_to_excel(data, filename="relatorio.xlsx"):
    try:
        # Converter dados para DataFrame
        df = pd.DataFrame(data['usuarios'])
        
        # Formatar cabeçalhos garantindo substituição de underscores
        df.columns = [col.replace('_', ' ').title() for col in df.columns]
        df.columns = [col.replace('-', ' ').title() for col in df.columns]
        
        # Criar workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Adicionar cabeçalhos formatados
        for col_num, column_title in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = column_title.replace('_', ' ')  # Garantir substituição
            cell.value = column_title.replace('-', ' ')
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
            cell.font = Font(bold=True, color='FFFFFF', size=12)
        
        # Adicionar dados
        for r_idx, row in enumerate(df.values, 2):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        
        wb.save(filename)
        return filename
        
    except Exception as e:
        st.error(f"Erro ao salvar planilha: {e}")
        return None

# Função para extrair colunas numéricas
def extract_numeric_columns(data):
    """Extrai colunas numéricas do DataFrame"""
    if not isinstance(data, pd.DataFrame):
        return None
    
    return data.select_dtypes(include=['int64', 'float64']).columns.tolist()

# Configura interface com Streamlit
st.title("Assistente de Criação de Planilhas")
st.write("Descreva a planilha que deseja criar e eu cuidarei do resto!")

def analyze_numeric_data(data):
    """Analisa dados numéricos do DataFrame"""
    try:
        df = pd.DataFrame(data['usuarios'])
        numeric_cols = extract_numeric_columns(df)
        
        if not numeric_cols:
            st.warning("Nenhuma coluna numérica encontrada na planilha")
            return
            
        col_to_analyze = st.selectbox(
            "Selecione a coluna para análise",
            numeric_cols
        )
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("### Análise Básica")
            st.metric("Soma", f"{df[col_to_analyze].sum():.2f}")
            st.metric("Média", f"{df[col_to_analyze].mean():.2f}")
            st.metric("Máximo", f"{df[col_to_analyze].max():.2f}")
            st.metric("Mínimo", f"{df[col_to_analyze].min():.2f}")
        
        with col2:
            st.write("### Estatísticas")
            st.metric("Mediana", f"{df[col_to_analyze].median():.2f}")
            st.metric("Desvio Padrão", f"{df[col_to_analyze].std():.2f}")
            st.metric("Variância", f"{df[col_to_analyze].var():.2f}")
        
        st.write("### Distribuição")
        fig = px.histogram(df, x=col_to_analyze, title=f"Distribuição de {col_to_analyze}")
        st.plotly_chart(fig, use_container_width=True)
        
    except Exception as e:
        st.error(f"Erro na análise: {str(e)}")

if user_input := st.text_area("Descreva a planilha desejada:", placeholder="Exemplo: Crie uma tabela com nomes, idades e cidades"):
    if st.button("Gerar Planilha"):
        with st.spinner("Processando..."):
            try:
                use_yes_no = st.session_state.use_yes_no
                data = process_prompt_to_data(user_input, use_yes_no)
                
                if data:
                    st.session_state.df = pd.DataFrame(data['usuarios'])
                    st.session_state.calculos_aplicados = False
                    st.success("Planilha gerada com sucesso!")
            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")

# Modificar a seção de cálculos e download
if st.session_state.df is not None:
    st.write("### Preview dos Dados Originais:")
    st.dataframe(st.session_state.df)
    
    st.write("### Opções de Download")
    download_option = st.radio(
        "Escolha uma opção:",
        ["Baixar planilha original", "Aplicar cálculos e baixar"]
    )
    
    if download_option == "Baixar planilha original":
        if st.button("Baixar Planilha Original"):
            filename = save_data_to_excel({'usuarios': st.session_state.df.to_dict('records')})
            if filename:
                with open(filename, "rb") as f:
                    bytes_data = f.read()
                st.download_button(
                    "Download Excel",
                    bytes_data,
                    file_name="planilha_original.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        colunas_numericas = st.session_state.df.select_dtypes(include=['int64', 'float64']).columns
        if len(colunas_numericas) > 0:
            coluna_calculo = st.selectbox("Selecione a coluna:", colunas_numericas)
            calculos_selecionados = st.multiselect("Selecione os cálculos:", list(CALCULOS.keys()))
            
            if calculos_selecionados:
                # Criar cópia para não modificar os dados originais
                df_calculado = aplicar_calculos(st.session_state.df.copy(), coluna_calculo, calculos_selecionados)
                
                # Mostrar preview dos dados calculados
                st.write("### Preview dos Dados com Cálculos:")
                st.dataframe(df_calculado)
                
                # Botão para download
                if st.button("Baixar Planilha com Cálculos"):
                    filename = save_data_to_excel({'usuarios': df_calculado.to_dict('records')})
                    if filename:
                        with open(filename, "rb") as f:
                            bytes_data = f.read()
                        st.download_button(
                            "Download Excel com Cálculos",
                            bytes_data,
                            file_name="planilha_com_calculos.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )