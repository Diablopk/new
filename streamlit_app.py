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
    "LLama 3.3 70B Versatile": "llama-3.3-70b-versatile",
    "Mixtral 8x7B": "mixtral-8x7b-32768",
    "Gemma 2 9B": "gemma2-9b-it",
    "LLama 3.1 8B Instant": "llama-3.1-8b-instant",
    "LLama 3 70B": "llama3-70b-8192",
    "LLama 3 8B": "llama3-8b-8192",
    
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
if 'color_rules' not in st.session_state:
    st.session_state.color_rules = {'valores': {}, 'colunas': {}}

# Adicionar após a inicialização do session_state
if 'edit_rules' not in st.session_state:
    st.session_state.edit_rules = {
        'deleted_cells': set(),  # Armazenar células excluídas (row, col)
        'centered_cells': set(),  # Armazenar células centralizadas
        'edited_cells': {}  # Armazenar texto editado {(row, col): novo_texto}
    }

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
    """Aplica os cálculos selecionados na coluna usando apenas nome da função"""
    for calculo in calculos_selecionados:
        if calculo in CALCULOS:
            # Usar apenas o nome do cálculo como nova coluna
            nova_coluna = f"{calculo}"
            df[nova_coluna] = CALCULOS[calculo](df[coluna])
    return df

def process_prompt_to_data(prompt, use_yes_no):
    try:
        formatted_prompt = f"""
        Gere um JSON válido contendo uma lista de objetos para uma tabela.
        O JSON deve ter apenas uma chave principal contendo um array de objetos.
        Exemplo: {{"dados": [{{"coluna1": "valor1", "coluna2": ["item1", "item2"]}}]}}
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
        
        # Encontrar primeira chave com array
        main_key = next((k for k, v in data.items() if isinstance(v, list)), None)
        
        if main_key:
            normalized_data = {"usuarios": data[main_key]}
            
            # Aplicar conversão Sim/Não se necessário
            if use_yes_no:
                normalized_data = convert_to_yes_no(normalized_data)
            else:
                normalized_data = normalize_data(normalized_data)
                
            return normalized_data
        else:
            st.error("Formato de dados inválido - necessário array de objetos")
            return None
            
    except Exception as e:
        st.error(f"Erro no processamento: {str(e)}")
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
    """Normaliza os dados mantendo formatação consistente"""
    try:
        if isinstance(data, dict):
            for key in data.keys():
                if isinstance(data[key], list):
                    normalized_items = []
                    for item in data[key]:
                        normalized_item = {}
                        for k, v in item.items():
                            # Formatar nome da coluna
                            col_name = k.replace('_', ' ').replace('-', ' ').title()
                            
                            if isinstance(v, list):
                                normalized_item[col_name] = ', '.join(str(x).replace('_', ' ').replace('-', ' ') for x in v)
                            else:
                                normalized_item[col_name] = v
                        normalized_items.append(normalized_item)
                    data[key] = normalized_items
        return data
    except Exception as e:
        st.error(f"Erro na normalização: {e}")
        return data

def convert_to_yes_no(data):
    """Converte listas em formato Sim/Não com formatação consistente"""
    if not isinstance(data, dict) or 'usuarios' not in data:
        return data
    
    modified_data = {'usuarios': []}
    
    for item in data['usuarios']:
        new_item = {}
        for k, v in item.items():
            if isinstance(v, list):
                col_prefix = k.replace('_', ' ').replace('-', ' ').title()
                for val in set(v):
                    new_key = f"{col_prefix} {str(val).replace('_', ' ').replace('-', ' ').title()}"
                    new_item[new_key] = 'Sim' if val in v else 'Não'
            else:
                new_item[k.replace('_', ' ').replace('-', ' ').title()] = v
        modified_data['usuarios'].append(new_item)
    
    return modified_data

def format_header(text):
    """Formata o texto substituindo _ e - por espaço"""
    return text.replace('_', ' ').replace('-', ' ').title()

def extract_color_hints(data):
    """Extrai cores especificadas no JSON de resposta"""
    color_hints = {}
    try:
        if isinstance(data, dict) and 'usuarios' in data:
            for item in data['usuarios']:
                for key, value in item.items():
                    if isinstance(value, str):
                        # Busca por padrão valor#COR
                        if '#' in value:
                            val, color = value.split('#', 1)
                            if len(color) == 6 and all(c in '0123456789ABCDEF' for c in color.upper()):
                                val = val.strip()
                                color_hints[val] = color.upper()
    except Exception as e:
        st.error(f"Erro ao extrair cores: {e}")
    return color_hints

def save_data_to_excel(data, filename="relatorio.xlsx"):
    try:
        # Extrair dicas de cores do JSON
        color_hints = extract_color_hints(data)
        
        # Converter dados para DataFrame, removendo os códigos de cor
        df = pd.DataFrame(data['usuarios'])
        for col in df.columns:
            df[col] = df[col].apply(lambda x: x.split('#')[0].strip() if isinstance(x, str) and '#' in x else x)
        
        # Formatar cabeçalhos
        df.columns = [format_header(col) for col in df.columns]
        
        # Criar workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Estilos padrão
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=12)
        cell_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        cell_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        
        # Adicionar cabeçalhos formatados
        for col_num, column_title in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = format_header(column_title)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = cell_border
            cell.alignment = cell_alignment
        
        # Adicionar dados com edições
        for r_idx, row in enumerate(df.values, 2):
            for c_idx, value in enumerate(row, 1):
                # Pular células excluídas
                if (r_idx-2, df.columns[c_idx-1]) in st.session_state.edit_rules['deleted_cells']:
                    continue
                
                cell = ws.cell(row=r_idx, column=c_idx)
                
                # Aplicar texto editado
                if (r_idx-2, df.columns[c_idx-1]) in st.session_state.edit_rules['edited_cells']:
                    cell.value = st.session_state.edit_rules['edited_cells'][(r_idx-2, df.columns[c_idx-1])]
                else:
                    cell.value = value
                
                # Aplicar centralização
                if (r_idx-2, df.columns[c_idx-1]) in st.session_state.edit_rules['centered_cells']:
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                else:
                    cell.alignment = cell_alignment
                
                cell.border = cell_border
                
                # Aplicar cor apenas se especificada no JSON
                str_value = str(value).strip()
                if str_value in color_hints:
                    cell.fill = PatternFill(
                        start_color=color_hints[str_value],
                        end_color=color_hints[str_value],
                        fill_type='solid'
                    )
                elif r_idx % 2 == 0:  # Manter zebrado para melhor visualização
                    cell.fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        
        # Ajustar largura das colunas
        for col_num in range(1, len(df.columns) + 1):
            col_letter = get_column_letter(col_num)
            max_length = 0
            
            # Verificar comprimento do cabeçalho
            header_cell = ws[f"{col_letter}1"]
            if header_cell.value:
                max_length = len(str(header_cell.value))
            
            # Verificar comprimento dos dados
            for cell in ws[col_letter][1:]:
                if cell.value:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
            
            # Definir largura da coluna (com margem de 2 caracteres)
            adjusted_width = max_length + 2
            ws.column_dimensions[col_letter].width = min(adjusted_width, 50)  # limita a 50 caracteres
        
        # Ajustar altura das linhas
        for row in ws.rows:
            max_height = 0
            for cell in row:
                if cell.value:
                    text_lines = str(cell.value).count('\n') + 1
                    estimated_height = text_lines * 15  # 15 pontos por linha
                    max_height = max(max_height, estimated_height)
            if max_height > 0:
                ws.row_dimensions[cell.row].height = max_height
        
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

# Modificar apenas a seção de interface após a geração da planilha
if st.session_state.df is not None:
    st.write("### Preview dos Dados Originais:")
    st.dataframe(st.session_state.df, key="df_preview_original")
    
    st.write("### Opções de Formatação")
    with st.expander("Personalizar Cores das Células"):
        st.write("Selecione valores e cores para destacar na planilha")
        
        # Opção de colorir valor específico ou coluna inteira
        modo_cor = st.radio(
            "Modo de coloração:",
            ["Valor específico", "Coluna inteira"],
            key="modo_cor_radio"
        )
        
        # Obter valores únicos de todas as colunas
        valores_unicos = {}
        for col in st.session_state.df.columns:
            valores_unicos[col] = st.session_state.df[col].unique().tolist()
        
        # Interface para adicionar cores
        col1, col2, col3 = st.columns(3)
        with col1:
            coluna_selecionada = st.selectbox(
                "Selecione a coluna:",
                list(valores_unicos.keys()),
                key="coluna_selector_formatacao"
            )
        with col2:
            if modo_cor == "Valor específico":
                valor_selecionado = st.selectbox(
                    "Selecione o valor:",
                    [str(v) for v in valores_unicos[coluna_selecionada]],
                    key="valor_selector_formatacao"
                )
        with col3:
            cor_selecionada = st.color_picker(
                "Escolha a cor:",
                "#FFFFFF",
                key="cor_picker_formatacao"
            )
        
        # Botão para adicionar cor
        if st.button("Adicionar Cor", key="btn_add_cor_formatacao"):
            if modo_cor == "Valor específico":
                st.session_state.color_rules['valores'][valor_selecionado] = cor_selecionada.replace('#', '')
                st.success(f"Cor aplicada ao valor: {valor_selecionado}")
            else:
                st.session_state.color_rules['colunas'][coluna_selecionada] = cor_selecionada.replace('#', '')
                st.success(f"Cor aplicada à coluna: {coluna_selecionada}")
        
        # Preview com cores aplicadas
        if st.session_state.color_rules['valores'] or st.session_state.color_rules['colunas']:
            st.write("### Preview com Cores:")
            df_preview = st.session_state.df.copy()
            
            def highlight_cells(x):
                styles = pd.DataFrame('', index=x.index, columns=x.columns)
                for valor, cor in st.session_state.color_rules['valores'].items():
                    for col in x.columns:
                        mask = x[col].astype(str) == valor
                        styles.loc[mask, col] = f'background-color: #{cor}'
                for coluna, cor in st.session_state.color_rules['colunas'].items():
                    styles[coluna] = f'background-color: #{cor}'
                return styles
            
            st.dataframe(df_preview.style.apply(highlight_cells, axis=None), key="df_preview_cores")
            
            if st.button("Limpar Todas as Cores", key="btn_limpar_cores_formatacao"):
                st.session_state.color_rules = {'valores': {}, 'colunas': {}}
                st.success("Todas as cores foram removidas")

    # Adicionar nova seção de edição
    st.write("### Opções de Edição")
    with st.expander("Editar Células"):
        st.write("Selecione células para editar, centralizar ou excluir")
        
        edit_tab1, edit_tab2, edit_tab3 = st.tabs(["Editar Texto", "Centralizar", "Excluir"])
        
        with edit_tab1:
            col1, col2, col3 = st.columns(3)
            with col1:
                edit_col = st.selectbox(
                    "Selecione a coluna:",
                    st.session_state.df.columns,
                    key="edit_col_selector"
                )
            with col2:
                edit_row = st.number_input(
                    "Selecione a linha (0 = todas):",
                    min_value=0,
                    max_value=len(st.session_state.df),
                    key="edit_row_input"
                )
            with col3:
                novo_texto = st.text_input(
                    "Novo texto:",
                    key="edit_text_input"
                )
            
            if st.button("Aplicar Edição", key="btn_apply_edit"):
                if edit_row == 0:  # Aplicar em toda a coluna
                    for idx in st.session_state.df.index:
                        st.session_state.edit_rules['edited_cells'][(idx, edit_col)] = novo_texto
                else:
                    st.session_state.edit_rules['edited_cells'][(edit_row-1, edit_col)] = novo_texto
                st.success("Texto editado com sucesso!")
        
        with edit_tab2:
            col1, col2 = st.columns(2)
            with col1:
                center_col = st.selectbox(
                    "Selecione a coluna:",
                    st.session_state.df.columns,
                    key="center_col_selector"
                )
            with col2:
                center_row = st.number_input(
                    "Selecione a linha (0 = todas):",
                    min_value=0,
                    max_value=len(st.session_state.df),
                    key="center_row_input"
                )
            
            if st.button("Centralizar", key="btn_center"):
                if center_row == 0:
                    for idx in st.session_state.df.index:
                        st.session_state.edit_rules['centered_cells'].add((idx, center_col))
                else:
                    st.session_state.edit_rules['centered_cells'].add((center_row-1, center_col))
                st.success("Células centralizadas!")
        
        with edit_tab3:
            col1, col2 = st.columns(2)
            with col1:
                del_col = st.selectbox(
                    "Selecione a coluna:",
                    st.session_state.df.columns,
                    key="del_col_selector"
                )
            with col2:
                del_row = st.number_input(
                    "Selecione a linha (0 = todas):",
                    min_value=0,
                    max_value=len(st.session_state.df),
                    key="del_row_input"
                )
            
            if st.button("Excluir Célula(s)", key="btn_delete"):
                if del_row == 0:
                    for idx in st.session_state.df.index:
                        st.session_state.edit_rules['deleted_cells'].add((idx, del_col))
                else:
                    st.session_state.edit_rules['deleted_cells'].add((del_row-1, del_col))
                st.success("Células excluídas!")
        
        if st.button("Limpar Todas as Edições", key="btn_clear_edits"):
            st.session_state.edit_rules = {
                'deleted_cells': set(),
                'centered_cells': set(),
                'edited_cells': {}
            }
            st.success("Todas as edições foram removidas!")

        # Preview com edições
        if any(rules for rules in st.session_state.edit_rules.values()):
            st.write("### Preview com Edições:")
            df_preview = st.session_state.df.copy()
            
            def highlight_edits(x):
                styles = pd.DataFrame('', index=x.index, columns=x.columns)
                
                for (row, col) in st.session_state.edit_rules['deleted_cells']:
                    styles.iloc[row, df_preview.columns.get_loc(col)] = 'background-color: #FFE6E6'
                
                for (row, col) in st.session_state.edit_rules['centered_cells']:
                    styles.iloc[row, df_preview.columns.get_loc(col)] = 'text-align: center'
                
                return styles
            
            # Aplicar edições de texto para preview
            for (row, col), texto in st.session_state.edit_rules['edited_cells'].items():
                df_preview.iloc[row, df_preview.columns.get_loc(col)] = texto
            
            st.dataframe(df_preview.style.apply(highlight_edits, axis=None), key="df_preview_edits")

    # Opções de Download - Remover seção duplicada e manter apenas esta
    st.write("### Opções de Download")
    download_option = st.radio(
        "Escolha uma opção:",
        ["Baixar planilha original", "Aplicar cálculos e baixar"],
        key="download_option_radio_unico"
    )
    
    if download_option == "Baixar planilha original":
        if st.button("Baixar Planilha"):
            # Criar cópia dos dados para aplicar cores
            data_with_colors = {'usuarios': st.session_state.df.to_dict('records')}
            
            # Aplicar cores definidas pelo usuário
            if 'color_rules' in st.session_state:
                for record in data_with_colors['usuarios']:
                    for key, value in record.items():
                        str_value = str(value)
                        # Aplicar cores por valor
                        if str_value in st.session_state.color_rules['valores']:
                            record[key] = f"{value}#{st.session_state.color_rules['valores'][str_value]}"
                        # Aplicar cores por coluna
                        elif key in st.session_state.color_rules['colunas']:
                            record[key] = f"{value}#{st.session_state.color_rules['colunas'][key]}"
            
            filename = save_data_to_excel(data_with_colors)
            if filename:
                with open(filename, "rb") as f:
                    bytes_data = f.read()
                st.download_button(
                    "Download Excel",
                    bytes_data,
                    file_name="planilha_formatada.xlsx",
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
