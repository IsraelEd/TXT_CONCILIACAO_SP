import re
import io
import unicodedata
import pandas as pd
import streamlit as st

# ==========================
# CONFIGURAÇÃO DA PÁGINA
# ==========================
st.set_page_config(page_title="Conciliação de Relatórios", layout="wide")
st.title("Conciliação de Relatórios — NF x Retenção")
st.caption("Upload de planilhas, normalização de dados e conciliação com tolerância fixa de R$ 1,00.")

TOLERANCIA = 1.00          # Fixo
CHAVE = ('NF', 'CNPJ')     # A chave é NF + CNPJ

# ==========================
# FUNÇÕES DE APOIO
# ==========================

def remover_acentos(s: str) -> str:
    s = str(s).replace('\n', ' ').replace('\r', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return ''.join(
        c for c in unicodedata.normalize('NFD', s)
        if unicodedata.category(c) != 'Mn'
    )

def localizar_coluna(df: pd.DataFrame, candidatos: list[str]) -> str:
    mapa = {remover_acentos(str(c).lower()): c for c in df.columns}
    # exato
    for cand in candidatos:
        chave = remover_acentos(str(cand).lower())
        for k, original in mapa.items():
            if k == chave:
                return original
    # contém
    for cand in candidatos:
        chave = remover_acentos(str(cand).lower())
        for k, original in mapa.items():
            if chave in k:
                return original
    colunas = ', '.join(map(str, df.columns))
    raise KeyError(f"Não encontrei nenhuma coluna entre {candidatos}. Colunas disponíveis: {colunas}")

def extrair_digitos(valor) -> str:
    """Extrai apenas os números, prevenindo o erro do Pandas transformar '73' em '73.0' e virar '730'"""
    if pd.isna(valor):
        return ""
    s = str(valor).strip()
    if s.lower() in ['nan', 'none', '<na>', 'nat', '']:
        return ""
    # Se o Pandas importou como float terminando em .0, removemos o .0 antes de limpar
    if s.endswith('.0'):
        s = s[:-2]
    return re.sub(r'\D', '', s)

def normalizar_nf(valor) -> str:
    return extrair_digitos(valor)

def normalizar_cnpj(valor) -> str:
    return extrair_digitos(valor)

def normalizar_filial(valor) -> str:
    digits = extrair_digitos(valor)
    if not digits:
        return ''
    return digits[-3:].zfill(3)

def parse_valor_brl(valor) -> float:
    if pd.isna(valor):
        return 0.0
    s = str(valor).strip()
    if s == '':
        return 0.0
    s = re.sub(r'[^\d,.\-]', '', s)
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s and '.' not in s:
        s = s.replace(',', '.')
    elif s.count('.') > 1 and ',' not in s:
        *milhares, dec = s.split('.')
        s = ''.join(milhares) + '.' + dec
    try:
        return float(s)
    except:
        return 0.0

def format_brl_str(x):
    """Formata número float em string PT-BR (ex.: 118,99)."""
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(x)

def format_date_br(series):
    """Converte série de datas para dd/mm/yyyy (string)."""
    s = pd.to_datetime(series, errors='coerce')
    return s.dt.strftime('%d/%m/%Y').fillna('')

def to_cnpjcpf_text(series):
    """Mantém apenas dígitos e retorna string formatada de forma blindada contra valores nulos."""
    def pad(val):
        d = extrair_digitos(val)
        if not d:
            return ""
        if len(d) <= 11:
            return d.zfill(11)  # CPF
        elif len(d) <= 14:
            return d.zfill(14)  # CNPJ
        else:
            return d            # não altera
    return series.apply(pad)

def selecionar_colunas_base(base, localizar_coluna):
    col_n_doc            = localizar_coluna(base, ['Nº documento', 'No documento', 'Numero documento'])
    col_lcto             = localizar_coluna(base, ['Lançamento Contábil', 'Lancamento Contabil'])
    col_data_doc         = localizar_coluna(base, ['Data documento'])
    col_data_lcto        = localizar_coluna(base, ['Data de lançamento', 'Data lancamento'])
    col_nfe              = localizar_coluna(base, ['Nº da nota fiscal eletrônica', 'No da nota fiscal eletronica', 'NF', 'NFe'])
    col_id_parc          = localizar_coluna(base, ['ID parceiro', 'ID Parceiro'])
    col_razao            = localizar_coluna(base, ['Razão Social', 'Razao Social'])
    col_dom_fiscal_parc  = localizar_coluna(base, ['Dom. Fiscal Parceiro', 'Dom Fiscal Parceiro'])
    col_desc_dom_fiscal  = localizar_coluna(base, ['Desc. Dom. Fiscal', 'Descricao Dom Fiscal', 'Desc Dom Fiscal'])
    col_filial           = localizar_coluna(base, ['Filial'])
    col_cnpjcpf          = localizar_coluna(base, ['CNPJCPF', 'CNPJ', 'CPF'])
    col_val_total_nota   = localizar_coluna(base, ['Valor Total da Nota'])
    col_base_calculo     = localizar_coluna(base, ['Base de Cálculo', 'Base de Calculo'])
    col_aliquota         = localizar_coluna(base, ['Alíquota', 'Aliquota'])
    col_imp_retido       = localizar_coluna(base, ['Imposto Retido', 'Retido'])
    try:
        col_imp_recolher = localizar_coluna(base, ['Imposto a Recolher'])
    except Exception:
        col_imp_recolher = None
    col_code_controle    = localizar_coluna(base, ['Code Controle', 'Código Controle', 'Cod Controle'])
    col_usuario          = localizar_coluna(base, ['Usuário', 'Usuario'])

    ordem_cols = [
        col_n_doc, col_lcto, col_data_doc, col_data_lcto, col_nfe, col_id_parc, col_razao,
        col_dom_fiscal_parc, col_desc_dom_fiscal, col_filial, col_cnpjcpf, col_val_total_nota,
        col_base_calculo, col_aliquota, col_imp_retido,
        (col_imp_recolher if col_imp_recolher else 'Imposto a Recolher'),
        col_code_controle, col_usuario
    ]

    return {
        'n_doc': col_n_doc, 'lcto': col_lcto, 'data_doc': col_data_doc, 'data_lcto': col_data_lcto,
        'nfe': col_nfe, 'id_parc': col_id_parc, 'razao': col_razao, 'dom_fiscal_parc': col_dom_fiscal_parc,
        'desc_dom_fiscal': col_desc_dom_fiscal, 'filial': col_filial, 'cnpjcpf': col_cnpjcpf,
        'val_total': col_val_total_nota, 'base_calc': col_base_calculo, 'aliquota': col_aliquota,
        'imp_retido': col_imp_retido, 'imp_recolher': col_imp_recolher,
        'code_controle': col_code_controle, 'usuario': col_usuario,
        'ordem_cols': ordem_cols
    }

def montar_conciliacao(df_base: pd.DataFrame, df_rel: pd.DataFrame, tolerancia: float, chave=('NF', 'CNPJ')):
    conc = pd.merge(
        df_base,
        df_rel,
        on=list(chave),
        how='outer',
        indicator=True,
        suffixes=('', '_rel')
    )

    conc['ImpostoRetido'] = conc['ImpostoRetido'].fillna(0.0)
    conc['Retencao']      = conc['Retencao'].fillna(0.0)
    conc['Diferenca']     = conc['ImpostoRetido'] - conc['Retencao']
    conc['AbsDiff']       = conc['Diferenca'].abs()

    def classificar_status(row):
        if row['_merge'] == 'left_only':
            return 'Lançamento indevido'
        elif row['_merge'] == 'right_only':
            return 'Nota provavelmente não lançada no SAP'
        else:
            return 'Conciliado' if row['AbsDiff'] <= tolerancia else 'Inconciliado'

    conc['Status'] = conc.apply(classificar_status, axis=1)

    inconc = conc[conc['Status'] == 'Inconciliado'].copy()
    indevido = conc[conc['Status'] == 'Lançamento indevido'].copy()
    nao_lancada = conc[conc['Status'] == 'Nota provavelmente não lançada no SAP'].copy()

    resumo_status = conc.groupby('Status', dropna=False).agg(
        Qtde=('NF', 'count'),
        Total_Base=('ImpostoRetido', 'sum'),
        Total_Rel=('Retencao', 'sum'),
        Soma_AbsDiff=('AbsDiff', 'sum')
    ).reset_index()

    totais = pd.DataFrame({
        'Indicador': ['Total lançados', 'Conciliados', 'Inconciliados', 'Lançamento indevido', 'Nota provavelmente não lançada no SAP'],
        'Valor': [
            len(conc),
            (conc['Status'] == 'Conciliado').sum(),
            (conc['Status'] == 'Inconciliado').sum(),
            (conc['Status'] == 'Lançamento indevido').sum(),
            (conc['Status'] == 'Nota provavelmente não lançada no SAP').sum()
        ]
    })

    return conc, inconc, indevido, nao_lancada, resumo_status, totais

def to_excel_bytes(dfs_dict: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', datetime_format='yyyy-mm-dd', date_format='yyyy-mm-dd') as writer:
        for sheet_name, df in dfs_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

def encontrar_duplicidades_por_nf(df: pd.DataFrame, coluna_nf: str) -> pd.DataFrame:
    dup = (
        df.assign(NF_norm=df[coluna_nf].apply(normalizar_nf))
          .groupby('NF_norm', as_index=False)
          .size()
    )
    dup = dup[dup['size'] > 1].rename(columns={'NF_norm': 'NF', 'size': 'Qtde'})
    return dup

def encontrar_duplicidades_por_nf_cnpj(df: pd.DataFrame, col_nf: str, col_cnpj: str) -> pd.DataFrame:
    dup = (
        df.assign(NF_norm=df[col_nf].apply(normalizar_nf),
                  CNPJ_norm=df[col_cnpj].apply(normalizar_cnpj))
          .groupby(['NF_norm', 'CNPJ_norm'], as_index=False)
          .size()
    )
    dup = dup[dup['size'] > 1].rename(columns={'NF_norm': 'NF', 'CNPJ_norm': 'CNPJ', 'size': 'Qtde'})
    return dup


# ==========================
# UPLOAD
# ==========================
col1, col2 = st.columns(2)
with col1:
    arquivo_base = st.file_uploader("Relatório ZDPFISC (xlsx)", type=["xlsx"])
with col2:
    arquivo_rel = st.file_uploader("Relatório Prefeitura (xlsx)", type=["xlsx"])

if arquivo_base and arquivo_rel:
    try:
        base = pd.read_excel(arquivo_base, engine='openpyxl')
        rel  = pd.read_excel(arquivo_rel,  engine='openpyxl')

        # Mapeamento robusto das colunas
        col_lcto   = localizar_coluna(base, ['Lançamento Contábil', 'Lancamento Contabil'])
        col_nfe    = localizar_coluna(base, ['Nº da nota fiscal eletrônica', 'No da nota fiscal eletronica', 'NF', 'NFe'])
        col_razao  = localizar_coluna(base, ['Razão Social', 'Razao Social'])
        col_cnpj_b = localizar_coluna(base, ['CNPJCPF', 'CNPJ', 'CPF'])
        col_filial = localizar_coluna(base, ['Filial'])
        col_retido = localizar_coluna(base, ['Imposto Retido', 'Retido'])

        col_nf_rel  = localizar_coluna(rel,  ['NF', 'NFe'])
        col_cnpj_r  = localizar_coluna(rel,  ['CNPJ'])
        col_val_rel = localizar_coluna(rel,  ['Retenção', 'Retencao', 'Valor Retenção', 'Valor Retencao', 'Valor'])

        # Normalização
        df_base = pd.DataFrame({
            'LancamentoContabil': base[col_lcto],
            'NF':                 base[col_nfe].apply(normalizar_nf),
            'RazaoSocial':        base[col_razao],
            'CNPJ':               base[col_cnpj_b].apply(normalizar_cnpj),
            'Filial':             base[col_filial].apply(normalizar_filial),
            'ImpostoRetido':      base[col_retido].apply(parse_valor_brl)
        })

        df_rel = pd.DataFrame({
            'NF':        rel[col_nf_rel].apply(normalizar_nf),
            'CNPJ':      rel[col_cnpj_r].apply(normalizar_cnpj),
            'Retencao':  rel[col_val_rel].apply(parse_valor_brl)
        })

        # ========================================================
        # CORREÇÃO AQUI: Remove linhas fantasmas/vazias do Excel!
        # ========================================================
        df_base = df_base[df_base['NF'] != ""]
        df_rel = df_rel[df_rel['NF'] != ""]

        # Conciliação por NF + CNPJ
        conc, inconc, indevido, nao_lancada, resumo_status, totais = montar_conciliacao(
            df_base, df_rel, tolerancia=TOLERANCIA, chave=CHAVE
        )

        # ==========================
        # Duplicidades
        # ==========================
        dups_base_nf      = encontrar_duplicidades_por_nf(df_base, 'NF')
        dups_rel_nf       = encontrar_duplicidades_por_nf(df_rel,  'NF')
        dups_base_nf_cnpj = encontrar_duplicidades_por_nf_cnpj(df_base, 'NF', 'CNPJ')
        dups_rel_nf_cnpj  = encontrar_duplicidades_por_nf_cnpj(df_rel,  'NF', 'CNPJ')

        qtd_dups_base_nf      = len(dups_base_nf)
        qtd_dups_rel_nf       = len(dups_rel_nf)
        qtd_dups_base_nf_cnpj = len(dups_base_nf_cnpj)
        qtd_dups_rel_nf_cnpj  = len(dups_rel_nf_cnpj)

        # ====== DASHBOARD ======
        st.subheader("Dashboard")
        colA, colB, colC, colD, colE = st.columns(5)
        colA.metric("Total lançados", len(conc))
        colB.metric("Conciliados", int((conc['Status'] == 'Conciliado').sum()))
        colC.metric("Inconciliados", int((conc['Status'] == 'Inconciliado').sum()))
        colD.metric("Indevidos", int((conc['Status'] == 'Lançamento indevido').sum()))
        colE.metric("Não lançados", int((conc['Status'] == 'Nota provavelmente não lançada no SAP').sum()))

        if any([qtd_dups_base_nf, qtd_dups_rel_nf, qtd_dups_base_nf_cnpj, qtd_dups_rel_nf_cnpj]):
            st.warning("Foram encontradas duplicidades. Verifique as tabelas abaixo.")

            if qtd_dups_base_nf > 0:
                with st.expander("Duplicidades no ZDPFISC por NF (mesma numeração aparecendo mais de 1x)"):
                    st.dataframe(dups_base_nf, use_container_width=True)
            if qtd_dups_rel_nf > 0:
                with st.expander("Duplicidades no Relatório da Prefeitura por NF"):
                    st.dataframe(dups_rel_nf, use_container_width=True)
            if qtd_dups_base_nf_cnpj > 0:
                with st.expander("Duplicidades no ZDPFISC por NF + CNPJ"):
                    st.dataframe(dups_base_nf_cnpj, use_container_width=True)
            if qtd_dups_rel_nf_cnpj > 0:
                with st.expander("Duplicidades no Relatório da Prefeitura por NF + CNPJ"):
                    st.dataframe(dups_rel_nf_cnpj, use_container_width=True)

        with st.expander("Resumo por Status"):
            st.dataframe(resumo_status, use_container_width=True)

        st.subheader("Conciliação geral")
        cols_conc = ['NF', 'RazaoSocial', 'LancamentoContabil', 'CNPJ', 'Filial','Diferenca', 'Status']
        conc_view = conc.loc[:, [c for c in cols_conc if c in conc.columns]]
        st.dataframe(
            conc_view,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Diferenca': st.column_config.NumberColumn('Diferença', format="%.2f")
            }
        )

        st.subheader("Inconciliados (diferença > R$ 1,00)")
        cols_inconc = ['NF', 'RazaoSocial', 'LancamentoContabil', 'CNPJ', 'Filial', 'Diferenca', 'Status']
        inconc_view = inconc.loc[:, [c for c in cols_inconc if c in inconc.columns]]
        st.dataframe(
            inconc_view,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Diferenca': st.column_config.NumberColumn('Diferença', format="%.2f")
            }
        )

        st.subheader("Lançamento indevido")
        if 'Diferenca' not in indevido.columns and {'ImpostoRetido','Retencao'}.issubset(indevido.columns):
            indevido['Diferenca'] = indevido['ImpostoRetido'] - indevido['Retencao']

        cols_fixas = ['NF', 'RazaoSocial', 'LancamentoContabil', 'CNPJ', 'Filial', 'Diferenca', 'Status']
        indevido_view = indevido.loc[:, [c for c in cols_fixas if c in indevido.columns]].copy()
        indevido_view.insert(0, 'Excluir', True)

        column_order = ['Excluir'] + [c for c in cols_fixas if c in indevido_view.columns]

        column_config = {
            'Excluir': st.column_config.CheckboxColumn(
                'Excluir', help="Quando marcado, a linha será excluída da ZDPFISC corrigida.", default=True
            ),
            'LancamentoContabil': st.column_config.TextColumn('LancamentoContábil', disabled=True),
            'NF':                 st.column_config.TextColumn('NF', disabled=True),
            'RazaoSocial':        st.column_config.TextColumn('RazaoSocial', disabled=True),
            'CNPJ':               st.column_config.TextColumn('CNPJ', disabled=True),
            'Filial':             st.column_config.TextColumn('Filial', disabled=True),
            'Diferenca':          st.column_config.NumberColumn('Diferença', disabled=True, format="%.2f"),
            'Status':             st.column_config.TextColumn('Status', disabled=True),
        }

        indevido_edit = st.data_editor(
            indevido_view,
            column_order=column_order,
            column_config=column_config,
            hide_index=True,
            use_container_width=True,
        )

        st.subheader("Nota que não consta no ZDPFISC")
        cols_nao_lancada = ['NF', 'CNPJ', 'Retencao', 'Status']
        nao_lancada_view = nao_lancada.loc[:, [c for c in cols_nao_lancada if c in nao_lancada.columns]]
        st.dataframe(nao_lancada_view, use_container_width=True, hide_index=True)

        # ==========================
        # ZDPFISC CORRIGIDO + LOG DE ALTERAÇÕES
        # ==========================
        st.markdown("---")
        st.subheader("Exportar ZDPFISC corrigido")

        m = selecionar_colunas_base(base, localizar_coluna)
        base_full = base.copy()
        base_full['_NF_key']   = base_full[m['nfe']].apply(normalizar_nf)
        base_full['_CNPJ_key'] = base_full[m['cnpjcpf']].apply(normalizar_cnpj)

        rel_keys = rel.copy()
        rel_keys['_NF_key']   = rel[col_nf_rel].apply(normalizar_nf)
        rel_keys['_CNPJ_key'] = rel[col_cnpj_r].apply(normalizar_cnpj)
        rel_keys = rel_keys[['_NF_key', '_CNPJ_key', col_val_rel]].rename(columns={col_val_rel: 'Retencao'})

        base_merge = pd.merge(
            base_full,
            rel_keys,
            on=['_NF_key', '_CNPJ_key'],
            how='left',
            indicator=True
        )

        left_only_df = base_merge[base_merge['_merge'] == 'left_only'].copy()
        decisao_indevido = indevido_edit[['NF', 'CNPJ', 'Excluir']].copy()

        # Pegamos EXATAMENTE o que o usuário quer excluir
        pares_remover = set(tuple(x) for x in decisao_indevido.loc[decisao_indevido['Excluir'] == True, ['NF', 'CNPJ']].values)
        pares_manter  = set(tuple(x) for x in decisao_indevido.loc[decisao_indevido['Excluir'] == False, ['NF', 'CNPJ']].values)

        # (5) Montar ZDPFISC corrigido: Manter as que deram match ('both') + TUDO do 'left_only' exceto o que foi marcado pra remover
        base_keep_both = base_merge[base_merge['_merge'] == 'both'].copy()
        
        mask_remover_left = left_only_df.apply(lambda r: (r['_NF_key'], r['_CNPJ_key']) in pares_remover, axis=1)
        base_keep_left = left_only_df[~mask_remover_left].copy()
        
        base_corrigida = pd.concat([base_keep_both, base_keep_left], ignore_index=True)

        imp_ret_col = m['imp_retido']
        imp_rec_col = m['imp_recolher'] if m['imp_recolher'] else 'Imposto a Recolher'
        if m['imp_recolher'] is None and imp_rec_col not in base_corrigida.columns:
            base_corrigida[imp_rec_col] = base_corrigida[imp_ret_col]

        base_corrigida['ImpostoRetido_num'] = base_corrigida[imp_ret_col].apply(parse_valor_brl)
        base_corrigida['AbsDiff'] = (base_corrigida['ImpostoRetido_num'] - base_corrigida['Retencao']).abs()

        mask_corrigir = base_corrigida['AbsDiff'] > TOLERANCIA
        antes_ret = base_corrigida.loc[mask_corrigir, imp_ret_col].copy()
        antes_rec = base_corrigida.loc[mask_corrigir, imp_rec_col].copy()

        base_corrigida.loc[mask_corrigir, imp_ret_col] = base_corrigida.loc[mask_corrigir, 'Retencao']
        base_corrigida.loc[mask_corrigir, imp_rec_col] = base_corrigida.loc[mask_corrigir, 'Retencao']

        base_corrigida[imp_ret_col] = base_corrigida[imp_ret_col].apply(format_brl_str)
        base_corrigida[imp_rec_col] = base_corrigida[imp_rec_col].apply(format_brl_str)

        try:
            base_corrigida[m['data_doc']]  = format_date_br(base_corrigida[m['data_doc']])
        except Exception: pass
        try:
            base_corrigida[m['data_lcto']] = format_date_br(base_corrigida[m['data_lcto']])
        except Exception: pass

        try:
            base_corrigida[m['cnpjcpf']] = to_cnpjcpf_text(base_corrigida[m['cnpjcpf']])
        except Exception: pass

        try:
            base_corrigida[m['filial']] = (
                base_corrigida[m['filial']].astype(str).str.strip()
                .str.replace(r'\D', '', regex=True).str[-3:].str.zfill(3)
            )
        except Exception: pass

        ordem_final = [c for c in m['ordem_cols'] if c in base_corrigida.columns]
        base_corrigida_out = base_corrigida[ordem_final].copy()

        if m['cnpjcpf'] in base_corrigida_out.columns:
            base_corrigida_out[m['cnpjcpf']] = base_corrigida_out[m['cnpjcpf']].apply(lambda x: f"'{x}" if x != "" else "")
        if m['filial'] in base_corrigida_out.columns:
            base_corrigida_out[m['filial']] = base_corrigida_out[m['filial']].apply(lambda x: f"'{x}" if x != "" else "")

        # Apenas os itens da planilha que de fato existem
        removidos = left_only_df[mask_remover_left][[m['nfe'], m['cnpjcpf']]].copy()
        removidos['Acao'] = 'Removido por decisão (no ZDPFISC, não encontrado no RELATÓRIO)'

        mask_manter_left = left_only_df.apply(lambda r: (r['_NF_key'], r['_CNPJ_key']) in pares_manter, axis=1)
        mantidos = left_only_df[mask_manter_left][[m['nfe'], m['cnpjcpf']]].copy()
        mantidos['Acao'] = 'Mantido por decisão (no ZDPFISC, não encontrado no RELATÓRIO)'

        atualizados = base_corrigida.loc[mask_corrigir, [m['nfe'], m['cnpjcpf']]].copy()
        atualizados['Imposto Retido (original)']     = antes_ret.values
        atualizados['Imposto a Recolher (original)'] = antes_rec.values
        atualizados['Imposto Retido (corrigido)']    = base_corrigida.loc[mask_corrigir, imp_ret_col].values
        atualizados['Imposto a Recolher (corrigido)']= base_corrigida.loc[mask_corrigir, imp_rec_col].values

        for df_log in (atualizados, removidos, mantidos):
            if m['cnpjcpf'] in df_log.columns:
                df_log[m['cnpjcpf']] = to_cnpjcpf_text(df_log[m['cnpjcpf']])
                df_log[m['cnpjcpf']] = df_log[m['cnpjcpf']].apply(lambda x: f"'{x}" if x != "" else "")

        for col_fmt in ['Imposto Retido (original)', 'Imposto a Recolher (original)', 'Imposto Retido (corrigido)', 'Imposto a Recolher (corrigido)']:
            if col_fmt in atualizados.columns:
                atualizados[col_fmt] = atualizados[col_fmt].apply(format_brl_str)

        excel_bytes_conc = to_excel_bytes({
            'Conciliacao': conc,
            'Inconciliados': inconc,
            'Lançamento indevido': indevido,
            'Nota nao lancada': nao_lancada,
            'Dashboard': resumo_status
        })
        excel_bytes_base_corr = to_excel_bytes({'Base corrigida': base_corrigida_out})
        excel_bytes_log = to_excel_bytes({
            'Atualizados': atualizados,
            'Removidos': removidos,
            'Mantidos por decisão': mantidos
        })

        col_b1, col_b2, col_b3 = st.columns(3)
        with col_b1:
            st.download_button("Baixar Excel de conciliação", data=excel_bytes_conc, file_name="conciliacao_relatorios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_b2:
            st.download_button("Baixar ZDPFISC corrigido", data=excel_bytes_base_corr, file_name="base_corrigida.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_b3:
            st.download_button("Baixar LOG de alterações", data=excel_bytes_log, file_name="log_alteracoes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Falha ao processar:")
        st.exception(e)

else:
    st.info("Faça o upload das duas planilhas para iniciar a conciliação.")