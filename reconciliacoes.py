import pandas as pd
import streamlit as st
import math
import ast
import numpy as np
import os
import time
from io import BytesIO
from itertools import combinations
from fuzzywuzzy import fuzz
from collections import Counter


def preprocess(df):
    df["Descrição"] = df["Descrição"].str.upper().str.replace(r"[^\w\s]", "", regex=True)
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True).dt.date
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").astype(float).round(2)
    df.dropna()
    return df


def add_conciliacao_completa(row, type, index, nota): #listas com registo integral de movimentos conciliados
    if type == 'cnt':
        add_ledger_row = dict(zip(st.session_state.ledger_cols, row.values.tolist()))
        add_ledger_row.update({"index":index})            
        add_ledger_row.update({"ordem":st.session_state.num_conciliacao})
        add_ledger_row.update({"nota":nota})
        st.session_state.ledger_conciliados += [add_ledger_row]

    elif type == 'banco':
        add_bank_row = dict(zip(st.session_state.bank_cols, row.values.tolist()))
        add_bank_row.update({"index":index})            
        add_bank_row.update({"ordem":st.session_state.num_conciliacao})
        add_bank_row.update({"nota":nota})
        st.session_state.bank_conciliados += [add_bank_row]


def remove_conciliacao_completa(ordens_movimento: list):
    st.session_state.ledger_conciliados = [item for item in st.session_state.ledger_conciliados if item["ordem"] not in ordens_movimento]

    st.session_state.bank_conciliados = [item for item in st.session_state.bank_conciliados if item["ordem"] not in ordens_movimento]


def return_conciliacao_completa(ordens_movimento: list):
    remove_ledger = [item for item in st.session_state.ledger_conciliados if item["ordem"] in ordens_movimento]

    remove_bank = [item for item in st.session_state.bank_conciliados if item["ordem"] in ordens_movimento]

    df_ledger_remove = pd.DataFrame(remove_ledger).set_index("index")
    df_ledger_remove.drop(["ordem", "nota"], axis=1, inplace=True)
    df_bank_remove = pd.DataFrame(remove_bank).set_index("index")
    df_bank_remove.drop(["ordem", "nota"], axis=1, inplace=True)

    st.session_state.ledger = pd.concat([st.session_state.ledger, df_ledger_remove])
    st.session_state.bank = pd.concat([st.session_state.bank, df_bank_remove])


def match_transactions(ledger, bank, previous_matched=None, match_total=100, match_parcial=80):
    if 'num_conciliacao' not in st.session_state:
        st.session_state.num_conciliacao = 1

    if previous_matched:
        matches = previous_matched
    else:
        matches = []

    bank['DateDiff'] = None
    bank_falta = bank

    for i, ledger_row in ledger.iterrows():
        ledger_amount = ledger_row["Valor"]
        ledger_date = ledger_row["Data"]
        ledger_desc = ledger_row["Descrição"]
        candidates_original = bank[
            (bank["Valor"]==ledger_amount)
        ]

        if candidates_original.empty:
            continue

        penalizacao_valores_iguais = 3
        score_threshold = 50

        if len(candidates_original) == 1:
            initial_score = score_threshold
        else:
            initial_score = score_threshold - (len(candidates_original) * penalizacao_valores_iguais)
        
        candidates = bank_falta[
            (bank_falta["Valor"]==ledger_amount)
        ]

        if candidates.empty:
            continue

        best_score = 0
        best_match = None

        for j, bank_row in candidates.iterrows():
            score = 0
            bank_date = candidates.loc[j, 'Data']
            candidates.loc[j, 'DateDiff'] = abs((ledger_date - bank_date).days)
            min_score_data = 15
            date_score = max(initial_score - candidates.loc[j, 'DateDiff'], min_score_data)
            bank_desc = bank_row["Descrição"]
            ceiling_description_score = 2
            description_score = fuzz.token_sort_ratio(ledger_desc, bank_desc) / ceiling_description_score
            score = date_score + description_score
            best_score = max(best_score, score)

            if best_score == score:
                best_match = j

        if best_score >= match_parcial and best_match is not None:
            
            if best_score >= match_total:
                nota_match = "Match total"
            else:
                nota_match = "Match parcial"

            matches.append({
                "Ordem": st.session_state.num_conciliacao,
                "OrdemMovimento": i,
                "DataMovimento": ledger_date.strftime("%Y-%m-%d"),
                "DescricaoMovimento": ledger_desc,
                "ValorMovimento": float(ledger_amount),
                "DataExtrato": bank_falta.loc[best_match, "Data"].strftime("%Y-%m-%d"),
                "OrdemExtrato": best_match,
                "DescricaoBanco": bank_falta.loc[best_match, "Descrição"],
                "ValorBanco": float(bank_falta.loc[best_match, "Valor"]),
                "Score": int(best_score),
                "Nota": nota_match
            })

            add_conciliacao_completa(type='cnt', row=ledger_row, index=i, nota=nota_match)
            add_conciliacao_completa(type='banco', row=bank_row, index=j, nota=nota_match)

            st.session_state.num_conciliacao += 1

            bank_falta.drop([best_match], inplace=True)
            ledger.drop([i], inplace=True)

    return [matches, ledger, bank_falta]


def dataframes_to_excel(ledger_conciliados, bank_conciliados,  ledger: pd.DataFrame, bank: pd.DataFrame):

    if isinstance(ledger_conciliados, list) and len(ledger_conciliados) == 0:
        st.markdown("Não foram efetuadas quaisquer conciliações")
        return

    match_cnt = pd.DataFrame(ledger_conciliados).drop(["index", "Ver"], axis=1)
    match_banco = pd.DataFrame(bank_conciliados).drop(["index", "Ver"], axis=1)
    sheets = ['MATCH CNT', 'MATCH BANCO', 'CNT', 'BANCO']
    df_ledger = ledger.copy()
    df_ledger.drop(columns=["Ver"], inplace=True)

    df_bank = bank.copy()
    df_bank.drop(columns=["Ver"], inplace=True)

    output = BytesIO()

    if 'max_ordem' in st.session_state:
        match_cnt['ordem'] = match_cnt['ordem'] + st.session_state.max_ordem 
        match_banco['ordem'] = match_banco['ordem'] + st.session_state.max_ordem 
        match_cnt = pd.concat([st.session_state.pre_conciliados_cnt, match_cnt])
        match_banco = pd.concat([st.session_state.pre_conciliados_banco, match_banco])
        match_cnt['Data'] = pd.to_datetime(match_cnt['Data']).dt.date
        match_banco['Data'] = pd.to_datetime(match_banco['Data']).dt.date

    dataframes = [match_cnt, match_banco, df_ledger, df_bank]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for i, frame in enumerate(dataframes):
            frame.to_excel(writer, sheet_name=sheets[i], index=False)
    
    excel_bytes = output.getvalue()
    return excel_bytes


def download_file():
    excel_bytes = dataframes_to_excel(st.session_state.ledger_conciliados, st.session_state.bank_conciliados, st.session_state.ledger, st.session_state.bank)
    return excel_bytes


def testar_combos(df, valor, max_subset_size=2, cancel_tol=1e-4, versao=None):
    df_inicio = df[['Valor']]
    valor_x_2 = valor * 2

    if valor > 0:
        df_clean = df_inicio[ (df_inicio['Valor'] <= (valor_x_2)) & (df_inicio['Valor'] >= (-valor_x_2))]
    else:
        df_clean = df_inicio[ (df_inicio['Valor'] >= (valor_x_2)) & (df_inicio['Valor'] <= (-valor_x_2))]

    rows = list(df_clean.itertuples(index=True, name=None))
    n = len(rows)
    num_combos = sum(math.comb(n, r) for r in range(1, max_subset_size+1))
    if num_combos > 2_000_000:
        return num_combos

    vals = df_clean['Valor'].to_numpy()
    idxs = df_clean.index.to_numpy()

    comb_dict = {}
    for size in range(1, max_subset_size + 1):
        for comb in combinations(range(len(vals)), size):
            combo_idxs = tuple(idxs[list(comb)])
            combo_vals = tuple(vals[list(comb)])
            combo_sum = sum(combo_vals)
            comb_dict[combo_idxs] = {'values': combo_vals, 'sum': combo_sum}
    
    matching_combos = {k: v for k, v in comb_dict.items() if abs(v['sum'] - valor) <= cancel_tol}
    all_rows = []

    for i, combo_idxs in enumerate(matching_combos.keys(), start=1):
        matching_rows = df.loc[list(combo_idxs)].copy()
        matching_rows["#"] = i
        all_rows.append(matching_rows)

    if len(all_rows) > 0:
        df_matching = pd.concat(all_rows)
        return df_matching
    else:
        return 0


@st.dialog("Erro conciliação", on_dismiss="rerun")
def erro_conciliacao():
    st.error("A soma dos valores a conciliar não é equivalente")
    if st.button("Fechar"):
        st.session_state.show_error = False
    else:
        time.sleep(2)


def conciliar(ledger=None, bank=None, nota=None, dialog=True):
    if 'num_conciliacao' not in st.session_state:
        st.session_state.num_conciliacao = 1
    else:
        st.session_state.num_conciliacao += 1

    cnt_soma = round(ledger.loc[ledger['Ver']==True, 'Valor'].sum(),2)
    banco_soma = round(bank.loc[bank['Ver']==True, 'Valor'].sum(),2)

    if cnt_soma != banco_soma:
        st.session_state.show_error = True
        if dialog == True:
            erro_conciliacao()
            return
        else:
            return
    else:
        st.session_state.show_error = False

    ledger_date, ledger_amount, ledger_desc = [], [], []
    bank_date, bank_amount, bank_desc = [], [], []
    ledger_drop_idx = []
    bank_drop_idx = []

    if ledger is not None and len(ledger) > 0:
        for i, row in ledger.iterrows():
            if row['Ver'] is True:
                add_conciliacao_completa(type='cnt', row=row, index=i, nota=nota)
                ledger_date.append(row['Data'].strftime("%Y-%m-%d"))
                ledger_amount.append(row['Valor'])
                ledger_desc.append(row['Descrição'])
                if "Index" in ledger:
                    ledger_drop_idx.append(row['Index'])
                else:
                    ledger_drop_idx.append(i)

    if bank is not None and len(bank) > 0:
        for j, row in bank.iterrows():
            if row['Ver'] is True:
                add_conciliacao_completa(type='banco', row=row, index=j, nota=nota)
                bank_date.append(row['Data'].strftime("%Y-%m-%d"))
                bank_amount.append(row['Valor'])
                bank_desc.append(row['Descrição'])
                if "Index" in bank:
                    bank_drop_idx.append(row['Index'])
                else:    
                    bank_drop_idx.append(j)

    if ledger_drop_idx:
        st.session_state.ledger.drop(ledger_drop_idx, inplace=True)

    if bank_drop_idx:
        st.session_state.bank.drop(bank_drop_idx, inplace=True)

    if ledger_date or bank_date:
        st.session_state.conciliados.append({
            "Ordem": st.session_state.num_conciliacao,
            "OrdemMovimento": None,
            "DataMovimento": ledger_date,
            "DescricaoMovimento": ledger_desc,
            "ValorMovimento": ledger_amount,
            "DataExtrato": bank_date,
            "OrdemExtrato": None,
            "DescricaoBanco": bank_desc,
            "ValorBanco": bank_amount,
            "Nota": nota
        })

        st.session_state.edit_ledger = filtrar_df(st.session_state.ledger, st.session_state["Datas_Cnt"], st.session_state["Min_Valor_Cnt"], st.session_state["Max_Valor_Cnt"], st.session_state["Pesquisa_Cnt"])
        st.session_state.edit_bank = filtrar_df(st.session_state.bank, st.session_state["Datas_Banco"], st.session_state["Min_Valor_Banco"], st.session_state["Max_Valor_Banco"], st.session_state["Pesquisa_Banco"])


def filtrar_df(df, intervalo_datas, min_valor, max_valor, pesquisa):
    min_data, max_data = intervalo_datas
    min_data = pd.to_datetime(min_data).date()
    max_data = pd.to_datetime(max_data).date()
    termos_pesquisa = [term.strip() for term in pesquisa.split(";") if term.strip()]

    mask = df.copy(deep=True)

    if termos_pesquisa:
        mask = mask[
            (mask["Valor"].between(min_valor, max_valor)) &
            (mask["Data"].between(min_data, max_data)) &
            (mask["Descrição"].str.contains('|'.join(termos_pesquisa), case=False, na=False))
        ]

    else:
        mask = mask[
            (mask["Valor"].between(min_valor, max_valor)) &
            (mask["Data"].between(min_data, max_data))
        ]

    mask['Ver'] = False
    return mask


def combos_por_total(valor, max_subset_size, df_outra, original_df, outra_tbl):
    result_df = testar_combos(df_outra, valor, max_subset_size=max_subset_size, cancel_tol=1e-4, versao=None)
    if type(result_df) is not int:
        result_df['Tipo'] = outra_tbl
        result_df = pd.concat([original_df, result_df[['Tipo', '#', 'Data', 'Descrição', 'Valor', 'Ver']]
                            ])
    return result_df


def combos_por_documento(valor, max_subset_size, df_outra, row, outra_tbl):
    result_df = testar_combos(df_outra, valor, max_subset_size=max_subset_size, cancel_tol=1e-4, versao=None)
    if type(result_df) is not int:
        result_df['Tipo'] = outra_tbl
        result_df = pd.concat([row, result_df[['Tipo', '#', 'Data', 'Descrição', 'Valor', 'Ver']]
                            ])
    return result_df


def show_combos_por_tipo(fonte, tipo, max_subset_size = 2, periodo = False):
    if fonte == "cnt":
        df_outra = st.session_state.tbl_banco
        df = st.session_state.tbl_cnt
        outra_tbl = 'banco'
    else:
        df_outra = st.session_state.tbl_cnt
        df = st.session_state.tbl_banco
        outra_tbl = 'cnt'

    st.session_state.combo_clicked = 1
    original_df = df.copy()
    original_df = original_df[original_df['Ver']==True]
    original_df['#'] = "--"
    original_df['Tipo'] = fonte
    original_df['Ver'] = False

    df['Data'] = pd.to_datetime(df["Data"], errors="coerce")
    df_outra['Data'] = pd.to_datetime(df_outra["Data"], errors="coerce")
    
    if tipo == "Por total":
        df_all = pd.DataFrame()
        valor = round(df.loc[df['Ver']==True, 'Valor'].sum(),2)

        if periodo == 'dia':
            dias = df_outra['Data'].unique()
            for dia in dias:
                df_parcial = df_outra[df_outra['Data']==dia]
                result_parcial_df = combos_por_total(valor, max_subset_size, df_parcial, original_df, outra_tbl)

                if type(result_parcial_df) is not int:
                    df_all = pd.concat([df_all, result_parcial_df])
        
            if df_all.empty == False:
                st.session_state.result_df = df_all

        elif periodo == 'mes':
            anos = df_outra['Data'].dt.year.unique()

            for ano in anos:
                df_parcial_ano = df_outra[df_outra['Data'].dt.year == ano]
                meses = df_outra['Data'].dt.month.unique()

                for mes in meses:
                    df_parcial_mes = df_parcial_ano[df_parcial_ano['Data'].dt.month == mes]
                    result_parcial_df = combos_por_total(valor, max_subset_size, df_parcial_mes, original_df, outra_tbl)

                    if type(result_parcial_df) is not int:
                        df_all = pd.concat([df_all, result_parcial_df])

            if df_all.empty == False:
                st.session_state.result_df = df_all
                return st.session_state.result_df

        elif periodo == 'ano':
            anos = df_outra['Data'].dt.year.unique()

            for ano in anos:
                df_parcial = df_outra[pd.to_datetime(df_outra['Data']).dt.year == ano]
                result_parcial_df = combos_por_total(valor, max_subset_size, df_parcial, original_df, outra_tbl)

                if type(result_parcial_df) is not int:
                    df_all = pd.concat([df_all, result_parcial_df])

            if df_all.empty == False:
                st.session_state.result_df = df_all
                return st.session_state.result_df

        else:
            result_df = testar_combos(df_outra, valor, max_subset_size=max_subset_size, cancel_tol=1e-9, versao=None) 
            
            if type(result_df) is not int:
                result_df['Tipo'] = outra_tbl
                st.session_state.result_df = pd.concat([original_df, result_df[['Tipo', '#', 'Data', 'Descrição', 'Valor', 'Ver']]])
            else:
                st.session_state.result_df = result_df
                return st.session_state.result_df

    elif tipo == "Por documento":
        df_all = pd.DataFrame()
        for index, row in original_df.iterrows():
            valor = round(row['Valor'],2)
            row_original = row.to_frame().T

            if periodo == 'dia':
                dias = df_outra['Data'].unique()
                for dia in dias:
                    df_parcial = df_outra[df_outra['Data']==dia]
                    result_parcial_df = combos_por_documento(valor, max_subset_size, df_parcial, row_original, outra_tbl)

                    if type(result_parcial_df) is not int:
                        df_all = pd.concat([df_all, result_parcial_df])
            
                if df_all.empty == False:
                    st.session_state.result_df = df_all

            elif periodo == 'mes':
                anos = df_outra['Data'].dt.year.unique()

                for ano in anos:
                    df_parcial_ano = df_outra[df_outra['Data'].dt.year == ano]
                    meses = df_outra['Data'].dt.month.unique()

                    for mes in meses:
                        df_parcial_mes = df_parcial_ano[df_parcial_ano['Data'].dt.month == mes]
                        result_parcial_df = combos_por_documento(valor, max_subset_size, df_parcial_mes, row_original, outra_tbl)

                        if type(result_parcial_df) is not int:
                            df_all = pd.concat([df_all, result_parcial_df])

                if df_all.empty == False:
                    st.session_state.result_df = df_all

            elif periodo == 'ano':
                anos = df_outra['Data'].dt.year.unique()

                for ano in anos:
                    df_parcial = df_outra[df_outra['Data'].dt.year == ano]
                    result_parcial_df = combos_por_documento(valor, max_subset_size, df_parcial, row_original, outra_tbl)

                    if type(result_parcial_df) is not int:
                        df_all = pd.concat([df_all, result_parcial_df])

                if df_all.empty == False:
                    st.session_state.result_df = df_all

            else:
                result_df = testar_combos(df_outra, valor, max_subset_size=max_subset_size, cancel_tol=1e-9, versao=None)
                
                if type(result_df) is not int:
                    result_df['Tipo'] = outra_tbl
                    df_all = pd.concat([df_all, row_original, result_df[['Tipo', '#', 'Data', 'Descrição', 'Valor', 'Ver']]
                    ])

                if df_all.empty == False:
                    st.session_state.result_df = df_all
                else:
                    st.session_state.result_df = result_df

    return st.session_state.result_df


@st.dialog("Escolher opções", width="large")
def opcoes_combos(fonte):

    if 'show_error' not in st.session_state:
        st.session_state.show_error = False

    options = ["Por total", "Por documento"]

    col1, col2 = st.columns(2)
    with col1:
        escolher_opcao_tipo = st.selectbox("Escolher o tipo", options=options)
    with col2:
        escolher_max_doc = st.number_input("Num Documentos Máx", min_value=0, value=2)

    dia = st.checkbox("Agrupar apenas por dia")
    mes = st.checkbox("Agrupar apenas por mês")
    ano = st.checkbox("Agrupar apenas por ano")

    if dia == True:
        periodo = 'dia'
    elif mes == True:
        periodo = 'mes'
    elif ano == True:
        periodo = 'ano'
    else:
        periodo = None

    if st.button("Ver combinações"):
        with st.spinner("A executar"):
            st.session_state.select_all = 0
            st.session_state.result_df = 0
            st.session_state.combo_clicked = 1
            st.session_state.result_df = show_combos_por_tipo(fonte, escolher_opcao_tipo, max_subset_size = escolher_max_doc, periodo=periodo)

        if type(st.session_state.result_df) is int:
            del st.session_state['combo_clicked']
            if st.session_state.result_df == 0:
                return st.warning("Não foram encontradas combinações para os parâmetros selecionados")
            else:
                return st.error(f"Limite ultrapassado de 2,000,000 combinações. A pesquisa tem {st.session_state.result_df:,}.")

    if 'combo_clicked' in st.session_state:
        if 'result_df' in st.session_state and type(st.session_state.result_df) is not int:
            with st.container():
                if st.session_state.select_all == 1:
                    st.session_state.result_df['Ver'] = True
                elif 'select_all' not in st.session_state:
                    st.session_state.result_df['Ver'] = False
                else:
                    st.session_state.result_df['Ver'] = False

                matrix_combos = st.data_editor(data=st.session_state.result_df[['Tipo', '#', 'Data', 'Descrição', 'Valor', 'Ver']])

                nota = st.text_input("Nota: ", width='stretch')
                
                col1, col2 = st.columns(2)

                with col1:
                    combo_conciliar = st.button("Conciliar", width='stretch')
                
                with col2: 
                    select_all = st.button("Selecionar todos", width='stretch', help="Clickar 2x")


            if combo_conciliar:
                st.session_state.matrix_combos = matrix_combos
                if st.session_state.matrix_combos[(st.session_state.matrix_combos['Ver']==True) & (st.session_state.matrix_combos['Tipo']=="cnt")].index.has_duplicates or st.session_state.matrix_combos[(st.session_state.matrix_combos['Ver']==True) & (st.session_state.matrix_combos['Tipo']=="banco")].index.has_duplicates:
                    st.error("Não podem haver linhas duplicadas a conciliar")

                    lista_cnt = st.session_state.matrix_combos[st.session_state.matrix_combos['Tipo']=="cnt"]
                    duplicados_cnt = lista_cnt[lista_cnt.index.duplicated()].index.tolist()

                    lista_banco = st.session_state.matrix_combos[st.session_state.matrix_combos['Tipo']=="banco"]
                    duplicados_banco = lista_banco[lista_banco.index.duplicated()].index.tolist()

                    st.markdown(f"ID duplicados cnt: {duplicados_cnt}")
                    st.markdown(f"ID duplicados banco: {duplicados_banco}")

                else:
                    exec_combo_conciliar(nota=nota)
                    if st.session_state.show_error == True:
                        st.error("A soma dos valores a conciliar não é equivalente")
                    else:
                        del st.session_state['combo_clicked']
                        del st.session_state['result_df']
                        st.session_state.dialog_close = 1
                        st.rerun()

            if select_all:
                st.session_state.select_all = 1


def exec_combo_conciliar(nota=None):
    ledger_combo = st.session_state.matrix_combos[(st.session_state.matrix_combos['Ver']==True) & (st.session_state.matrix_combos['Tipo']=='cnt')]
    bank_combo = st.session_state.matrix_combos[(st.session_state.matrix_combos['Ver']==True) & (st.session_state.matrix_combos['Tipo']=='banco')]

    ledger_combo = ledger_combo.index.tolist()
    bank_combo = bank_combo.index.tolist()

    ledger_combo = st.session_state.ledger[st.session_state.ledger.index.isin(ledger_combo)].copy()
    ledger_combo['Ver'] = True

    bank_combo = st.session_state.bank[st.session_state.bank.index.isin(bank_combo)].copy()
    bank_combo  ['Ver'] = True

    conciliar(ledger=ledger_combo, bank=bank_combo, nota=nota, dialog=False)
    del st.session_state["matrix_combos"]


def manual(ledger, bank):

    if ledger.empty == False:
        min_data_cnt = ledger["Data"].min()
        max_data_cnt = ledger["Data"].max()
        min_valor_cnt = ledger["Valor"].min()
        max_valor_cnt = ledger["Valor"].max()
    else:
        min_data_cnt = '1900-01-01'
        max_data_cnt = '1900-01-01'
        min_valor_cnt = 0
        max_valor_cnt = 0

    if bank.empty == False:
        min_data_banco = bank["Data"].min()
        max_data_banco= bank["Data"].max()
        min_valor_banco = bank["Valor"].min()
        max_valor_banco = bank["Valor"].max()
    else:
        min_data_banco = '1900-01-01'
        max_data_banco = '1900-01-01'
        min_valor_banco = 0
        max_valor_banco = 0

    if 'dialog_close' not in st.session_state:
        st.session_state.dialog_close = 1

    with st.container():
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.subheader("Contabilidade")
            intervalo_datas_cnt = st.date_input("Datas", ( min_data_cnt, max_data_cnt ), min_data_cnt, max_data_cnt, key="Datas_Cnt", format="DD-MM-YYYY")
            pesquisa_cnt = st.text_input("Pesquisa: separar termos com ';'", key="Pesquisa_Cnt")
        with col2:
            st.subheader(" ")
            input_min_valor_cnt = st.number_input(label="Valor mínimo", min_value=min_valor_cnt, max_value=max_valor_cnt, value=min_valor_cnt, 
            key="Min_Valor_Cnt")
            input_max_valor_cnt = st.number_input(label="Valor máximo", min_value=min_valor_cnt, max_value=max_valor_cnt, value=max_valor_cnt, key="Max_Valor_Cnt")
        with col3:
            st.subheader("Banco")
            intervalo_datas_banco = st.date_input("Datas", ( min_data_banco, max_data_banco ), min_data_banco, max_data_banco, key="Datas_Banco", format='DD-MM-YYYY')
            pesquisa_banco = st.text_input("Pesquisa: separar termos com ';'", key="Pesquisa_Banco")
        with col4:
            st.subheader(" ")
            input_min_valor_banco = st.number_input(label="Valor mínimo", min_value=min_valor_banco, max_value=max_valor_banco, value=min_valor_banco, 
            key="Min_Valor_Banco")
            input_max_valor_banco = st.number_input(label="Valor máximo", min_value=min_valor_banco, max_value=max_valor_banco, value=max_valor_banco, key="Max_Valor_Banco")

        col1a, col2a, col3a, col4a, col5a, col6a, col7a, col8a = st.columns([1,1,1,3,1,1,1,3])

        with col1a: filtrar_cnt = st.button("Atualizar", key="filtrar_cnt", width="stretch")
        with col2a: select_all_cnt = st.button("Selecionar", key="select_all_cnt", width="stretch")
        with col3a: deselect_all_cnt = st.button("Remover", key="deselect_all_cnt", width="stretch")
        with col4a: combos_cnt = st.button("Combinações possíveis", key="combos_cnt", width="stretch")
        with col5a: filtrar_banco = st.button("Atualizar", key="filtrar_banco", width="stretch")
        with col6a: select_all_banco = st.button("Selecionar", key="select_all_banco", width="stretch")
        with col7a: deselect_all_banco = st.button("Remover", key="deselect_all_banco", width="stretch")
        with col8a: combos_banco = st.button("Combinações possíveis", key="combos_banco", width="stretch")

        if filtrar_cnt or 'edit_ledger' not in st.session_state or st.session_state.dialog_close == 1:

            st.session_state.edit_ledger = filtrar_df(ledger, st.session_state["Datas_Cnt"], st.session_state["Min_Valor_Cnt"], st.session_state["Max_Valor_Cnt"], st.session_state["Pesquisa_Cnt"])
            st.session_state.dialog_close = 0
        
        if filtrar_banco or 'edit_bank' not in st.session_state  or st.session_state.dialog_close == 1:
            
            st.session_state.edit_bank = filtrar_df(bank, st.session_state["Datas_Banco"], st.session_state["Min_Valor_Banco"], st.session_state["Max_Valor_Banco"], st.session_state["Pesquisa_Banco"])
            st.session_state.dialog_close = 0

        if select_all_cnt:
            st.session_state.edit_ledger['Ver'] = True
            st.session_state.tbl_cnt['Ver'] = True
            st.rerun()

        if deselect_all_cnt:
            st.session_state.edit_ledger['Ver'] = False
            st.session_state.tbl_cnt['Ver'] = False
            st.rerun()

        if select_all_banco:
            st.session_state.edit_bank['Ver'] = True
            st.rerun()

        if deselect_all_banco:
            st.session_state.edit_bank['Ver'] = False
            st.rerun()

    px_per_row = 35



    with st.container():
        col1, col2 = st.columns([2,2])

        with col1:
            st.session_state.tbl_cnt = st.data_editor(
                data= st.session_state.edit_ledger,
                use_container_width=True,
                height= min(35 + (px_per_row * len(st.session_state.edit_ledger)), 500),
                hide_index=True,
                key='tbl_edit_ledger'
            )

            soma_cnt = round(st.session_state.tbl_cnt['Valor'].sum(),2)
            st.markdown(f"Total {soma_cnt}")

        with col2:
            st.session_state.tbl_banco = st.data_editor(
                data=st.session_state.edit_bank,
                use_container_width=True,
                hide_index=True,
                height = min(35 + (px_per_row * len(st.session_state.edit_bank)), 500),
                key='tbl_edit_bank'
            )
            soma_banco = round(st.session_state.tbl_banco['Valor'].sum(),2)
            st.markdown(f"Total {soma_banco}")

        linhas_cnt = st.session_state.tbl_cnt[st.session_state.tbl_cnt['Ver']==True]
        soma_cnt = round(linhas_cnt["Valor"].sum(),2)
        linhas_bank = st.session_state.tbl_banco[st.session_state.tbl_banco['Ver']==True]
        soma_bank = round(linhas_bank["Valor"].sum(),2)
        dif_soma = round(soma_cnt - soma_bank,2)
        opcoes_cnt = len(st.session_state.tbl_cnt.loc[(st.session_state.tbl_cnt['Valor'] == -dif_soma)])
        opcoes_banco = len(st.session_state.tbl_banco.loc[(st.session_state.tbl_banco['Valor'] == dif_soma)])

        if dif_soma != 0 and opcoes_cnt == 0 and opcoes_banco == 0:
            st.error(f"Contabilidade: {soma_cnt} | Banco: {soma_bank} | Diferença: {dif_soma}") 
        elif dif_soma == 0:
            st.success(f"Contabilidade: {soma_cnt} | Banco: {soma_bank} | Diferença: 0")
        else:
            st.warning(f"Contabilidade: {soma_cnt} | Banco: {soma_bank} | Diferença: {dif_soma} ---- Existem {opcoes_cnt} opções na contabilidade com esta diferença --- Existem {opcoes_banco} opções no banco com esta diferença")

    if combos_cnt:
        if 'combo_clicked' in st.session_state:
            del st.session_state['combo_clicked']
        if 'result_df' in st.session_state:
            del st.session_state['result_df']
        opcoes_combos("cnt")

    if combos_banco:
        if 'combo_clicked' in st.session_state:
            del st.session_state['combo_clicked']
        if 'result_df' in st.session_state:
            del st.session_state['result_df']
        opcoes_combos("banco")

    nota = st.text_input("Nota", key=f"nota_conciliacao")
    fazer_match = st.button("Conciliar", key=f'conciliar_bnt')

    if fazer_match: 
        conciliar(st.session_state.tbl_cnt, st.session_state.tbl_banco, nota)
        st.rerun()

    st.divider()
    
    if st.button("Preparar ficheiro download"):
        excel_bytes = download_file()
        st.download_button(
            label="Download ficheiro",
            data=excel_bytes,
            file_name='reconciliacao.xlsx',
            mime='application/vnd.ms-excel'
        )


def reconciliacao_inicial(ficheiro):
    ledger_df = ficheiro["CNT"]
    bank_df = ficheiro["BANCO"]
    ledger_df = preprocess(ledger_df)
    bank_df = preprocess(bank_df)

    st.session_state.ledger_cols = ledger_df.columns.values.tolist()
    st.session_state.bank_cols = bank_df.columns.values.tolist()

    if st.session_state.conciliar == True:
        conciliados, ledger, bank = match_transactions(ledger=ledger_df, bank=bank_df, match_total= st.session_state.match_total, match_parcial=st.session_state.match_parcial)
        st.session_state.ledger = ledger
        st.session_state.bank = bank
        st.session_state.conciliados = conciliados
    else:
        st.session_state.ledger = ledger_df
        st.session_state.bank = bank_df
        st.session_state.conciliados = {
                "Ordem": 0,
                "OrdemMovimento": 0,
                "DataMovimento": None,
                "DescricaoMovimento": None,
                "ValorMovimento": 0,
                "DataExtrato": None,
                "OrdemExtrato": 0,
                "DescricaoBanco": None,
                "ValorBanco": 0,
                "Score": None,
                "Nota":None
            }

    st.session_state.conciliar = False


def remover_conciliacao(tbl): #remove linha conciliada da lista de conciliacoes e devolve aos pendentes
    col_cnt = ['DataMovimento', 'DescricaoMovimento', 'ValorMovimento']
    col_bnk = ['DataExtrato', 'DescricaoBanco', 'ValorBanco']
    
    tbl_processada = pd.DataFrame(tbl)
    linhas_a_remover = tbl_processada[tbl_processada['Remover']==True]
    linhas_a_remover_list = linhas_a_remover["Ordem"].values.tolist()

    return_conciliacao_completa(ordens_movimento=linhas_a_remover_list)
    remove_conciliacao_completa(ordens_movimento=linhas_a_remover_list)

    st.session_state.edit_ledger = st.session_state.ledger
    st.session_state.edit_bank = st.session_state.bank

    idx_a_remover = linhas_a_remover['Ordem'].to_list()

    st.session_state.conciliados[:] = [
    linha for linha in st.session_state.conciliados if linha.get("Ordem") not in idx_a_remover
    ]

    st.session_state.remove_success = 1

    st.rerun()


def adicionar_linhas(df, type):
    if type == 'cnt':
        pd.concat(st.session_state.ledger, df)

    elif type == 'banco':
        pd.concat(st.session_state.bank, df)


def conciliados():

    with st.form("conciliados"):
        lista_keys = ['Ordem', 'DataMovimento', 'DescricaoMovimento','ValorMovimento', 'DataExtrato', 'DescricaoBanco', 'ValorBanco', 'Nota']
        conciliados_show =  [
            {k: d[k] for k in lista_keys if k in d}
            for d in st.session_state.conciliados
        ]

        for row in conciliados_show:
            row['Remover'] = False

        tbl_conciliados = st.data_editor(
            data=conciliados_show
        )

        if st.form_submit_button("Remover"):
            remover_conciliacao(tbl_conciliados)

        if 'remove_success' not in st.session_state:
            st.session_state.remove_success = 0

        if st.session_state.remove_success == 1:
            st.success("Movimentos removidos com sucesso")
            st.session_state.remove_success = 0


def resumo():
    total_cnt = 0
    total_banco = 0

    for row in st.session_state.conciliados:
        val_cnt = row.get("ValorMovimento", 0)
        if isinstance(val_cnt, list): 
            total_cnt += sum(val_cnt)
        else:                      
            total_cnt += val_cnt

        val_banco = row.get("ValorBanco", 0)
        if isinstance(val_banco, list):  
            total_banco += sum(val_banco)
        else:                      
            total_banco += val_banco

    dados = [st.session_state.ledger, st.session_state.bank]
    num_docs = []
    valor_por_conciliar = []

    for info in dados:
        contagem = 0
        valor = 0
        for i, row in info.iterrows():
            val = row["Valor"]
            if isinstance(val, list):
                contagem += len(val)
                valor += val.sum()
            else:                      
                contagem += 1
                valor += val
        num_docs.append(contagem)
        valor_por_conciliar.append(valor)

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        st.subheader("Saldo Inicial")
        st.metric(label="Contabilidade", value=f'{st.session_state.saldo_i_cnt} €')
        st.metric(label="Banco", value=f'{st.session_state.saldo_i_banco} €')
    with col2:
        st.subheader("Contabilidade")
        st.metric(label="Por conciliar", value=f'{round(valor_por_conciliar[0],2)} €')
        st.metric(label="Por conciliar (lançamentos)", value=num_docs[0])
    with col3:
        st.subheader("Banco")
        st.metric(label="Por conciliar", value=f'{round(valor_por_conciliar[1],2)} €')
        st.metric(label="Por conciliar (lançamentos)", value=num_docs[1])
    with col4:
        st.subheader("Diferenças")
        st.metric(label="Diferença", value=f'{round((st.session_state.saldo_i_banco - st.session_state.saldo_i_cnt + valor_por_conciliar[0] - valor_por_conciliar[1]),2)} €')


def app():
    st.set_page_config(page_title="Reconciliações", page_icon=':star', layout='wide')

    if 'conciliacao_inicial' not in st.session_state:
        st.session_state.conciliacao_inicial = None
        st.session_state.conciliar = True

    if 'file' not in st.session_state:
        st.session_state.file = None

    if st.session_state.conciliacao_inicial == None:
        pag_inicial = st.empty()
        with pag_inicial.container():
            st.title("Reconciliação inicial")
            col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)
            with col1:
                st.session_state.saldo_i_banco = st.number_input("Saldo Inicial Banco", value=0.00)
            with col2:
                st.session_state.saldo_i_cnt = st.number_input("Saldo Inicial Contabilidade", value=0.00)
            with col3:
                st.session_state.match_total = st.number_input("MatchTotal%", value=100)
            with col4:
                st.session_state.match_parcial = st.number_input("MatchParcial%", value=80)

            st.divider()
            col1, col2 = st.columns(2)

            with st.spinner("A processar..."):

                if 'ledger_conciliados' not in st.session_state:
                    st.session_state.ledger_conciliados = []
                    st.session_state.bank_conciliados = []

                with col1:
                    st.session_state.file = st.file_uploader("Fazer upload do ficheiro. Deve ter 2 folhas com nome CNT e BANCO cada uma com as colunas Data, Descrição e Valor. Podem existir outras colunas", type=["xlsx", "xls"])

                    if st.button("Executar programa"):
                        if st.session_state.file == None:
                            st.error("Não foi efetuado upload do ficheiro")
                        else:
                            xls = pd.ExcelFile(st.session_state.file)
                            dfs = {}
                            for sheet in ['CNT', 'BANCO']:
                                df = pd.read_excel(xls, sheet_name=sheet)
                                df = df.dropna(how='all').dropna(axis=1, how='all')
                                dfs[sheet] = df

                            if ('MATCH CNT' in xls.sheet_names) and ('MATCH BANCO' in xls.sheet_names):
                                st.session_state.pre_conciliados_cnt = pd.read_excel(xls, sheet_name='MATCH CNT')
                                st.session_state.pre_conciliados_banco = pd.read_excel(xls, sheet_name='MATCH BANCO')
                                st.session_state.max_ordem = st.session_state.pre_conciliados_cnt['ordem'].max()

                            st.session_state.file = dfs
                            reconciliacao_inicial(st.session_state.file)
                            st.session_state.conciliacao_inicial = 1
                            st.rerun()

                    st.divider()

                    if st.button("Demonstração"):
                        file = os.path.join("file", "demo.xlsx")
                        st.session_state.file = pd.read_excel(file, sheet_name=None)
                        reconciliacao_inicial(st.session_state.file)
                        st.session_state.conciliacao_inicial = 1
                        st.rerun()

                    path_template = os.path.join("file", "template.xlsx")

                    with open(path_template, "rb") as f:
                        st.download_button(
                            label="Download Template",
                            data=f,
                            file_name="template.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

    elif st.session_state.conciliacao_inicial == 1:
        del st.session_state['file']
        st.session_state.bank['Ver'] = False
        st.session_state.ledger['Ver'] = False

        colunas_prioritarias = ['Data', 'Descrição', 'Valor', 'Ver']

        colunas_prioritarias_cnt = [col for col in colunas_prioritarias if col in st.session_state.ledger.columns]
        colunas_prioritarias_banco = [col for col in colunas_prioritarias if col in st.session_state.bank.columns]

        outras_colunas_cnt = [col for col in st.session_state.ledger.columns if col not in colunas_prioritarias_cnt]
        outras_colunas_banco = [col for col in st.session_state.bank.columns if col not in colunas_prioritarias_banco]

        st.session_state.ledger = st.session_state.ledger[colunas_prioritarias + outras_colunas_cnt]
        st.session_state.bank = st.session_state.bank[colunas_prioritarias + outras_colunas_banco]

        st.session_state.ledger = st.session_state.ledger.sort_values(by="Data")
        st.session_state.bank = st.session_state.bank.sort_values(by="Data")

        st.session_state.ledger_cols = st.session_state.ledger.columns
        st.session_state.bank_cols = st.session_state.bank.columns

        if 'DateDiff' in st.session_state.bank.columns:
            st.session_state.bank.drop(columns=['DateDiff'], inplace=True)

        with st.container():
            opcao = st.selectbox("", options=['Resumo', 'Conciliados', 'Manual'])

            if opcao == 'Conciliados':
                conciliados()
            elif opcao == 'Manual':
                manual(st.session_state.ledger, st.session_state.bank)
            elif opcao == "Resumo":
                resumo()


if __name__ == '__main__':
    app()