import streamlit as st
import pandas as pd
import requests
import time
import unicodedata
import os
import io
import re
import pdfplumber
import urllib3

# Desativa avisos de SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==== CABEÇALHO / RODAPÉ (constantes e helpers) ====
HEADER_TIPO_REGISTRO = "1003"  # 4 caracteres fixos
FOOTER_TIPO_REGISTRO = "9"     # 1 caractere fixo

def montar_cabecalho(ccm8: str, data_inicial_yyyymmdd: str, data_final_yyyymmdd: str) -> str:
    """
    Cabeçalho: 1003(4) + CCM(8) + DATA_INICIAL(8) + DATA_FINAL(8) = 28 caracteres.
    Exemplo: 1003803242902025100120251031
    """
    head = HEADER_TIPO_REGISTRO + str(ccm8).zfill(8) + str(data_inicial_yyyymmdd).zfill(8) + str(data_final_yyyymmdd).zfill(8)
    if len(head) != 28:
        raise ValueError(f"Cabeçalho inválido: {len(head)} (esperado 28)")
    return head

def somar_coluna_float(df: pd.DataFrame, col: str) -> float:
    """
    Soma uma coluna da planilha usando a função converter_float (mantém convenção dos seus valores).
    """
    total = 0.0
    if col in df.columns:
        for v in df[col]:
            total += converter_float(v)
    return total


def somar_deducoes(df: pd.DataFrame) -> float:
    """
    Soma as deduções da planilha linha a linha, usando a regra:
    Dedução = max(0, Valor Total da Nota - Base de Cálculo)
    (usa converter_float para respeitar o formato da planilha)
    """
    total_ded = 0.0
    col_vt = 'Valor Total da Nota'
    col_bc = 'Base de Cálculo'
    if col_vt in df.columns and col_bc in df.columns:
        for _, r in df.iterrows():
            vt = converter_float(r.get(col_vt, 0))
            bc = converter_float(r.get(col_bc, 0))
            ded = max(0.0, vt - bc)
            total_ded += ded
    return total_ded


def montar_rodape(qtd_notas: int, valor_total_soma: float, base_calculo_soma: float) -> str:
    """
    Rodapé: '9'(1) + QTDE(7) + VALOR_TOTAL(15) + BASE_CALCULO(15) = 38 caracteres.
    Campos de 15 posições formatados com formatar_valor (centavos, zero-fill).
    Exemplo: 90000061000001509417927000000000000000
    """
    foot = (
        FOOTER_TIPO_REGISTRO +
        str(int(qtd_notas)).zfill(7) +
        formatar_valor(valor_total_soma, 15) +
        formatar_valor(base_calculo_soma, 15)
    )
    if len(foot) != 38:
        raise ValueError(f"Rodapé inválido: {len(foot)} (esperado 38)")
    return foot



# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador TXT ISS", layout="wide")

# --- FUNÇÕES AUXILIARES ---

def remover_acentos(texto):
    """Remove acentos e converte para maiúsculas (ASCII puro)."""
    if not isinstance(texto, str):
        return str(texto) if texto is not None else ""
    nfkd = unicodedata.normalize('NFKD', texto)
    texto_sem_acento = u"".join([c for c in nfkd if not unicodedata.combining(c)])
    return texto_sem_acento.upper().encode('ascii', 'ignore').decode('ascii')

def converter_float(valor):
    """Converte valores do Excel para float."""
    if pd.isna(valor) or valor == "":
        return 0.0
    if isinstance(valor, (float, int)):
        return float(valor)
    valor_str = str(valor).strip()
    try:
        # 4.500,00 -> 4500.00
        if '.' in valor_str and ',' in valor_str:
            valor_str = valor_str.replace('.', '').replace(',', '.')
        elif ',' in valor_str:
            valor_str = valor_str.replace(',', '.')
        return float(valor_str)
    except:
        return 0.0

def formatar_valor(valor_float, tamanho):
    """Formata float para string de zeros sem ponto."""
    try:
        valor_int = int(round(valor_float * 100))
        return str(valor_int).zfill(tamanho)
    except:
        return "0" * tamanho

def consultar_cnpj_api(cnpj):
    """Consulta BrasilAPI."""
    cnpj_limpo = "".join(filter(str.isdigit, str(cnpj)))
    if len(cnpj_limpo) != 14: return None
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_limpo}"
    try:
        response = requests.get(url, timeout=10, verify=False)
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 429:
            time.sleep(2)
            return consultar_cnpj_api(cnpj)
    except:
        pass
    return None

def ler_pdf_surems(caminho_arquivo):
    """Extrai códigos do PDF."""
    pares = []
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    matches = re.findall(r"(?m)^\s*(\d{4,5})\s+(\d{1,2}\.\d{1,2})", text)
                    pares.extend(matches)
    except Exception as e:
        st.error(f"Erro PDF {os.path.basename(caminho_arquivo)}: {e}")
    return pares

def carregar_tabela_surems():
    mapa = {}
    pasta = "bases_sp"
    if not os.path.exists(pasta): return mapa
    arquivos = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
    for arquivo in arquivos:
        for cod_surems, item_lei in ler_pdf_surems(os.path.join(pasta, arquivo)):
            chave = item_lei.replace('.', '').replace(',', '').strip().zfill(4)
            val = cod_surems.strip()
            if len(val) == 5 and val.startswith('0'): val = val[1:]
            mapa[chave] = val.zfill(4)
    return mapa


def normalizar_filial(valor):
    # Converte para string e remove espaços
    s = str(valor).strip()
    # Mantém só dígitos
    digits = ''.join(ch for ch in s if ch.isdigit())
    if not digits:
        return ''  # ou '000' se quiser um padrão quando vier vazio
    # Pega os 3 últimos dígitos e completa com zeros à esquerda
    return digits[-3:].zfill(3)


# --- APP ---

st.title("Gerador de TXT São Paulo")

st.sidebar.header("Configurações")


def ddmmaaaa_para_yyyymmdd(s: str) -> str:
    """
    Converte uma string 'ddmmaaaa' (apenas dígitos) para 'yyyymmdd'.
    Se inválida (tamanho != 8 ou contém não dígitos), retorna '00000000'.
    """
    if not s:
        return "00000000"
    s = "".join(ch for ch in str(s) if ch.isdigit())
    if len(s) != 8:
        return "00000000"
    dd = s[0:2]
    mm = s[2:4]
    aaaa = s[4:8]
    return f"{aaaa}{mm}{dd}"



# ==== CAMPOS ADICIONAIS PARA O TXT (sem prints/infos) ====
ccm = st.sidebar.text_input("CCM (8 caracteres)", value="", max_chars=8)
data_inicial_ddmmaaaa = st.sidebar.text_input("Data Inicial (DDMMAAAA)", value="", max_chars=8)
data_final_ddmmaaaa   = st.sidebar.text_input("Data Final (DDMMAAAA)", value="", max_chars=8)


uploaded_file = st.file_uploader("Carregar Planilha (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype=object)
    mapa_surems = carregar_tabela_surems()
    
    st.write(f"Linhas: {len(df)}")
    
    
    if st.button("Gerar Arquivo TXT"):
        output = io.StringIO()
        progress = st.progress(0)
        status = st.empty()
        erros = []
        


        # ==== CABEÇALHO NO TXT (primeira linha) — usando ddmmaaaa -> yyyymmdd ====
        header_data_inicial = ddmmaaaa_para_yyyymmdd(data_inicial_ddmmaaaa)
        header_data_final   = ddmmaaaa_para_yyyymmdd(data_final_ddmmaaaa)

        # Se quiser fallback automático (opcional): quando usuário deixar em branco, usar min/max da coluna 'Data documento'
        if header_data_inicial == "00000000" or header_data_final == "00000000":
            if "Data documento" in df.columns:
                serie_datas = pd.to_datetime(df["Data documento"].astype(str), dayfirst=True, errors="coerce")
                dt_min = serie_datas.min()
                dt_max = serie_datas.max()
                if header_data_inicial == "00000000" and pd.notna(dt_min):
                    header_data_inicial = dt_min.strftime("%Y%m%d")
                if header_data_final == "00000000" and pd.notna(dt_max):
                    header_data_final = dt_max.strftime("%Y%m%d")

        header_str = montar_cabecalho(ccm, header_data_inicial, header_data_final)

        # Mantém seu fluxo: output já criado acima no botão (se não, crie aqui)
        # output = io.StringIO()  # somente se não existir a linha no seu botão
        output.write(header_str + "\n")



        for i, row in df.iterrows():
            try:
                # --- DADOS ---
                
                filial = normalizar_filial(row.get('Filial', ''))

                nota = str(row.get('Nº da nota fiscal eletrônica', '')).strip()
                data_raw = pd.to_datetime(row.get('Data documento', ''), errors='coerce')
                
                v_total = converter_float(row.get('Valor Total da Nota', 0))
                v_base = converter_float(row.get('Base de Cálculo', 0))
                v_aliq = converter_float(row.get('Alíquota', 0))
                if v_aliq < 1 and v_aliq > 0: v_aliq *= 100
                
                retido_txt = str(row.get('Imposto Retido', '')).upper()
                cnpj = str(row.get('CNPJ/CPF', '')).strip()
                razao = str(row.get('Razão Social', '')).strip()
                lei_raw = str(row.get('Code Controle', '')).strip()
                
                # --- API ---
                dados_api = consultar_cnpj_api(cnpj)
                logradouro = numero = bairro = municipio = uf = cep = ""
                is_simples = False
                if dados_api:
                    logradouro = dados_api.get('logradouro', '')
                    numero = dados_api.get('numero', '')
                    bairro = dados_api.get('bairro', '')
                    municipio = dados_api.get('municipio', '')
                    uf = dados_api.get('uf', '')
                    cep = str(dados_api.get('cep', '')).replace('-', '').replace('.', '')
                    is_simples = bool(dados_api.get('opcao_pelo_simples', False))

                # --- CÁLCULO DEDUÇÃO (REVERTIDO) ---
                # Dedução = Valor Total - Base de Cálculo
                if v_base > 0 and v_base < v_total:
                    v_deducao = v_total - v_base
                else:
                    v_deducao = 0.0

                # --- MONTAGEM BLOCOS ---
                b1 = filial.ljust(8)
                b2 = "".join(filter(str.isdigit, nota)).zfill(9) + "   "
                
                # Bloco 3
                b3_data = data_raw.strftime("%Y%m%d") + "NT" if pd.notna(data_raw) else "00000000NT"
                b3_valor = formatar_valor(v_total, 15)
                
                # Campo Dedução calculado acima
                b3_campo_deducao = formatar_valor(v_deducao, 15)
                
                cod_lei = "".join(filter(str.isdigit, lei_raw)).zfill(4)
                cod_sur = mapa_surems.get(cod_lei, "0000")
                b3_aliq = str(int(v_aliq * 100)).zfill(3)
                b3_retido = "1" if any(x in retido_txt for x in ["SIM", "YES"]) else "2"
                b3_tipo = "2" if len("".join(filter(str.isdigit, cnpj))) > 11 else "1"
                b3_cnpj = "".join(filter(str.isdigit, cnpj)).zfill(14)
                
                separador = "0"
                
                b3_content = (
                    b3_data + 
                    b3_valor + 
                    b3_campo_deducao + 
                    separador +         # Separador Zero Obrigatório
                    cod_sur + 
                    cod_lei + 
                    "0" + 
                    b3_aliq + 
                    b3_retido + 
                    b3_tipo + 
                    b3_cnpj
                )
                b3 = b3_content + (" " * 8)
                
                # Endereço
                b4 = remover_acentos(razao).ljust(78)[:78]
                b5 = remover_acentos(logradouro).ljust(50)[:50]
                b6 = remover_acentos(numero).ljust(40)[:40]
                b7 = remover_acentos(bairro).ljust(30)[:30]
                b8 = remover_acentos(municipio).ljust(50)[:50]
                b9 = remover_acentos(f"{uf}{cep}").ljust(85)[:85]
                
                # B10 - Simples (Apenas os 2 dígitos, sem espaço ainda)
                cod_simples = "14" if is_simples else "10"
                
                # --- MONTAGEM LINHA FINAL COM TAMANHO ESTRITO ---
                
                # 1. Concatena tudo sem o espaço final
                raw_line = b1 + b2 + b3 + b4 + b5 + b6 + b7 + b8 + b9 + cod_simples
                
                # 2. Força o preenchimento ou corte exato em 432 caracteres
                line_content = raw_line.ljust(432)[:432]
                
                # 3. Adiciona o ÚNICO espaço final para fechar em 433
                line_final = line_content
                
                output.write(line_final + "\n")
                
                progress.progress((i + 1) / len(df))
                status.text(f"Processando {i+1}/{len(df)}")
                
            except Exception as e:
                erros.append(f"Linha {i+2}: {e}")

                

        # ==== RODAPÉ NO TXT (última linha) ====
        qtd_notas = len(df)
        valor_total_soma = somar_coluna_float(df, 'Valor Total da Nota')
        somdeducoes_soma = somar_deducoes(df)  #soma das deduções

        footer_str = montar_rodape(qtd_notas, valor_total_soma, somdeducoes_soma)

        # Rodapé no final do arquivo
        output.write(footer_str)

        st.success("Concluído!")
        st.download_button("Baixar TXT", output.getvalue(), "retencao_iss_final_fixed.txt", "text/plain")
        
        if erros:
            st.error("Erros:")
            st.write(erros)