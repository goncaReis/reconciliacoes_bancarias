import pandas as pd
import streamlit as st
import math
import ast
from io import BytesIO
from fuzzywuzzy import fuzz
from collections import Counter

def preprocess(df):
    df["Descrição"] = df["Descrição"].str.upper().str.replace(r"[^\w\s]", "", regex=True)
    df["Data"] = pd.to_datetime(df["Data"])
    return df


def match_transactions(ledger, bank, previous_matched=None, date_tolerance_days:int=365):
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
            (bank["Valor"]==ledger_amount) &
            (bank["Data"].between(ledger_date - pd.Timedelta(days=date_tolerance_days),
            ledger_date + pd.Timedelta(days=date_tolerance_days)))
        ]

        if candidates_original.empty:
            continue

        penalizacao_valores_iguais = 5
        initial_score = 80 - (len(candidates_original) * penalizacao_valores_iguais)
        candidates = bank_falta[
            (bank_falta["Valor"]==ledger_amount) &
            (bank_falta["Data"].between(ledger_date - pd.Timedelta(days=date_tolerance_days),
            ledger_date + pd.Timedelta(days=date_tolerance_days)))
        ]

        if candidates.empty:
            continue

        best_score = 0
        best_match = None

        for j, bank_row in candidates.iterrows():
            score = 0
            bank_date = candidates.loc[j, 'Data']
            candidates.loc[j, 'DateDiff'] = abs((ledger_date - bank_date).days)
            date_score = max(initial_score - candidates.loc[j, 'DateDiff'], 25)
            bank_desc = bank_row["Descrição"]
            description_score = fuzz.token_sort_ratio(ledger_desc, bank_desc) / 5
            score = date_score + description_score
            best_score = max(best_score, score)

            if best_score == score:
                best_match = j

        if best_score > 30 and best_match is not None:
            matches.append({
                "OrdemMovimento": i,
                "DataMovimento": ledger_date.strftime("%Y-%m-%d"),
                "DescricaoMovimento": ledger_desc,
                "ValorMovimento": ledger_amount,
                "DataExtrato": bank_falta.loc[best_match, "Data"].strftime("%Y-%m-%d"),
                "OrdemExtrato": best_match,
                "DescricaoBanco": bank_falta.loc[best_match, "Descrição"],
                "ValorBanco": bank_falta.loc[best_match, "Valor"],
                "Score": int(best_score),
                "Nota":"Automático"
            })
            
            bank_falta.drop([best_match], inplace=True)
            ledger.drop([i], inplace=True)

    return [matches, ledger, bank_falta]


def dataframes_to_excel(matches: list, ledger: pd.DataFrame, bank: pd.DataFrame):
    matched_df = pd.DataFrame(matches)
    sheets = ['MATCH', 'VALIDAR', 'FALTA']
    ledger['Tipo'] = 'Contabilidade'
    bank['Tipo'] = 'Banco'

    missing_df = pd.concat([ledger, bank], ignore_index=True)
    missing_df = missing_df[['Tipo', 'Data', 'Descrição', 'Valor']]
    missing_df.sort_values(by=['Valor', 'Data', 'Tipo'], inplace=True)

    evaluate_df = matched_df[matched_df['Score'] < 60]
    matched_df = matched_df[(matched_df['Score'] >= 60) | (matched_df['Score'].isna())]

    dataframes = [matched_df, evaluate_df, missing_df]

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for i, frame in enumerate(dataframes):
            frame.to_excel(writer, sheet_name=sheets[i], index=False)
    
    excel_bytes = output.getvalue()
    return excel_bytes


def upload_file(placeholder):
    with placeholder.container():
        st.title("Dados para reconcialiação um-para-um")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.session_state.saldo_i_banco = st.number_input("Saldo Inicial Banco", value=0.00)
        with col2:
            st.session_state.saldo_i_cnt = st.number_input("Saldo Inicial Contabilidade", value=0.00)
        with col3:
            tol_dias = st.number_input("Tolerância dias para conciliação um-para-um", value=365, min_value=0)
        st.divider()
        reconciliation_file = st.file_uploader("Fazer upload do ficheiro. Deve ter 2 folhas com nome CNT e BANCO cada uma com as colunas Data, Descrição e Valor.", type=["xlsx", "xls"])
        submit_button = st.button("Executar programa")
        if submit_button:
            if reconciliation_file == None:
                st.warning("Não foi efetuado upload do ficheiro")
            else:
                file = pd.read_excel(reconciliation_file, sheet_name=None)
                placeholder.empty()
                st.session_state.file = file
                st.session_state.days = tol_dias


def download_file():
    excel_bytes = dataframes_to_excel(st.session_state.conciliados, st.session_state.ledger, st.session_state.bank)
    return excel_bytes


def matrizes_manual(container):
    px_per_row = 35
    with container.container():
        col1, col2 = st.columns([2,2])
        if 'edit_ledger' in st.session_state:
            with col1:
                st.session_state.cnt = st.data_editor(
                    data=st.session_state.edit_ledger[['Data', 'Descrição', 'Valor','Ver']],
                    disabled=['Data', 'Descrição','Valor'],
                    use_container_width=True,
                    height = min(35 + (px_per_row * len(st.session_state.edit_ledger)), 500),
                    hide_index=True,
                    key=f'ledger_matrix_{st.session_state.version}'
                )
            with col2:
                st.session_state.bnk = st.data_editor(
                    data=st.session_state.edit_bank[['Data', 'Descrição','Valor', 'Ver']],
                    use_container_width=True,
                    hide_index=True,
                    height = min(35 + (px_per_row * len(st.session_state.edit_bank)), 500),
                    key=f'bank_matrix_{st.session_state.version}'
                )

        if 'cnt' in st.session_state:
            linhas_cnt = st.session_state.cnt[st.session_state.cnt['Ver']==True]
            soma_cnt = round(linhas_cnt["Valor"].sum(),2)
            linhas_bank = st.session_state.bnk[st.session_state.bnk['Ver']==True]
            soma_bank = round(linhas_bank["Valor"].sum(),2)
            dif_soma = round(soma_cnt - soma_bank,2)
            opcoes_cnt = len(st.session_state.cnt.loc[(st.session_state.cnt['Valor'] == -dif_soma)])
            opcoes_banco = len(st.session_state.bnk.loc[(st.session_state.bnk['Valor'] == dif_soma)])

            if dif_soma != 0 and opcoes_cnt == 0 and opcoes_banco == 0:
                st.error(f"Contabilidade: {soma_cnt} | Banco: {soma_bank} | Diferença: {dif_soma}") 
            elif dif_soma == 0:
                st.success(f"Contabilidade: {soma_cnt} | Banco: {soma_bank} | Diferença: 0")
            else:
                st.warning(f"Contabilidade: {soma_cnt} | Banco: {soma_bank} | Diferença: {dif_soma} ---- Existem {opcoes_cnt} opções na contabilidade com esta diferença --- Existem {opcoes_banco} opções no banco com esta diferença")


def handle_change_combos():
    st.session_state.combos = st.session_state.combos_editor


def testar_combos(df, valor, tolerancia_valor, max_subset_size=4, cancel_tol=1e-9, versao=None):
    combos = []
    df = df.drop(columns=["Ver"])
    rows = list(df.itertuples(index=True, name=None))
    n = len(rows)
    num_combos = sum(math.comb(n, r) for r in range(1, max_subset_size+1))

    if num_combos > 1000000:
        st.error(f"O número de combinações possíveis deve ser inferior a 1.000.000, atual = {num_combos}")
        return

    def search(start, depth, current_sum, current_rows, counter):
            if depth > max_subset_size:
                return

            # Check if current combination matches the target
            if depth > 0 and abs(current_sum - valor) <= tolerancia_valor:
                combos.append(list(current_rows))

            for i in range(start, n):
                row = rows[i]
                v = row[3]

                # Skip if this value cancels an existing one
                if v != 0 and counter.get(-v, 0) > 0:
                    continue

                # Branch-and-bound: remaining sum max
                remaining_sum_max = sum(abs(r[3]) for r in rows[i+1:])

                if abs(current_sum + v) - abs(valor) > tolerancia_valor + remaining_sum_max:
                    continue  # impossible to reach target

                # Include current row and recurse
                counter[v] += 1
                current_rows.append(row)
                search(i + 1, depth + 1, current_sum + v, current_rows, counter)
                current_rows.pop()
                counter[v] -= 1
                if counter[v] == 0:
                    del counter[v]

    search(0, 0, 0.0, [], Counter())

    if combos:
        flattened = []
        for i, combo in enumerate(combos, start=1):
            for row in combo:
                flattened.append({
                    "Index": row[0],
                    "Combo #": i,
                    "Data": row[2],
                    "Descrição": row[1],
                    "Valor": row[3],
                    "Ver": False
                })
        result_df = pd.DataFrame(flattened)
        return result_df     
    else:
        st.warning("Não existem combinações possíveis para os parametros introduzidos")


def ver_opcoes_reconciliacao(df):
    if df is st.session_state.ledger:
        df_outra = st.session_state.bank
        opcao_ledger = 0
    else:
        df_outra = st.session_state.ledger
        opcao_ledger = 1

    with st.container():
        min_data = min(st.session_state.ledger["Data"].min(), st.session_state.bank["Data"].min())
        max_data = max(st.session_state.ledger["Data"].max(), st.session_state.bank["Data"].max())
        min_valor = min(st.session_state.ledger["Valor"].min(), st.session_state.bank["Valor"].min())
        max_valor = max(st.session_state.ledger["Valor"].max(), st.session_state.bank["Valor"].max())

        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1: intervalo_datas = st.date_input("Datas", ( min_data, max_data ), min_data, max_data)
        with col2: pesquisa = st.text_input("Pesquisa")
        with col3: input_min_valor = st.number_input(label="Valor mínimo", min_value=min_valor, max_value=max_valor, value=min_valor)
        with col4: input_max_valor = st.number_input(label="Valor máximo", min_value=min_valor, max_value=max_valor, value=max_valor)
        with col5: max_size_combo = st.number_input(label="Num Docs Máx", value=4)

        col1, col2 = st.columns([2,2])
        with col1:
            tbl = st.data_editor(
                data=df[['Data', 'Descrição','Valor','Ver']],
                use_container_width=True,
                key='tabela_original'
            )
            nota = st.text_input("Nota")
            ver_combos_btn = st.button("Ver combinações")
            conciliar_clicked = st.button("Conciliar")

        with col2:
            if ver_combos_btn:
                if 'versao_combos' not in st.session_state:
                    st.session_state.versao_combos = 0
                else:
                    previous = f'result_combos_{st.session_state.versao_combos-1}'
                    if previous in st.session_state:
                        del st.session_state[previous]

                st.session_state.combos = None
                valor = tbl[tbl['Ver'] == True]['Valor'].sum()
                df_mask = filtrar_df(df_outra, intervalo_datas, input_min_valor, input_max_valor, pesquisa)
                result_df = testar_combos(valor=valor, df=df_mask, versao=st.session_state.versao_combos, max_subset_size=max_size_combo)
                st.session_state.combo_df = result_df
                st.session_state.versao_combos += 1

            if 'combo_df' in st.session_state:
                combo_df = st.data_editor(st.session_state.combo_df, hide_index=True, key=f"result_combos_{st.session_state.versao_combos}")

    if conciliar_clicked:
        if opcao_ledger == 0:
            conciliar(bank=combo_df, ledger=tbl[tbl['Ver']==True], nota=nota)
        elif opcao_ledger == 1:
            conciliar(ledger=combo_df, bank=tbl[tbl['Ver']==True], nota=nota)
        st.session_state.combo_df = None
        st.rerun()

    if st.button("Preparar ficheiro download"):
        excel_bytes = download_file()
        st.download_button(
            label="Download ficheiro",
            data=excel_bytes,
            file_name='reconciliacao.xlsx',
            mime='application/vnd.ms-excel'
        )


def conciliar(ledger=None, bank=None, nota=None):
    if 'num_conciliacao' not in st.session_state:
        st.session_state.num_conciliacao = 0
    else:
        st.session_state.num_conciliacao += 1

    ledger_date, ledger_amount, ledger_desc = [], [], []
    bank_date, bank_amount, bank_desc = [], [], []
    ledger_drop_idx = []
    bank_drop_idx = []

    if ledger is not None and len(ledger) > 0:
        for i, row in ledger.iterrows():
            if row['Ver'] is True:
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
            "OrdemMovimento": f'C{st.session_state.num_conciliacao}',
            "DataMovimento": ledger_date,
            "DescricaoMovimento": ledger_desc,
            "ValorMovimento": ledger_amount,
            "DataExtrato": bank_date,
            "OrdemExtrato": None,
            "DescricaoBanco": bank_desc,
            "ValorBanco": bank_amount,
            "Nota": nota
        })


def filtrar_df(df, intervalo_datas, min_valor, max_valor, pesquisa):
    min_data, max_data = intervalo_datas
    min_data = pd.to_datetime(min_data)
    max_data = pd.to_datetime(max_data)
    termos_pesquisa = [term.strip() for term in pesquisa.split(";") if term.strip()]

    if termos_pesquisa:
        mask = df[
            (df["Valor"].between(min_valor, max_valor)) &
            (df["Data"].between(min_data, max_data)) &
            (df["Descrição"].str.contains('|'.join(termos_pesquisa), case=False, na=False))
        ]
        return mask
    else:
        mask = df[
            (df["Valor"].between(min_valor, max_valor)) &
            (df["Data"].between(min_data, max_data))
        ]
        return mask


def manual(ledger, bank):
    min_data = min(ledger["Data"].min(), bank["Data"].min())
    max_data = max(ledger["Data"].max(), bank["Data"].max())
    min_valor = min(ledger["Valor"].min(), bank["Valor"].min())
    max_valor = max(ledger["Valor"].max(), bank["Valor"].max())

    if 'version' not in st.session_state:
        st.session_state.version = 0

    inputs_container = st.empty()

    with inputs_container.container():
        col1, col2, col3, col4 = st.columns(4)
        with col1: intervalo_datas = st.date_input("Datas", ( min_data, max_data ), min_data, max_data, key="Datas_Man")
        with col2: pesquisa = st.text_input("Pesquisa: separar termos com ';'", key="Pesquisa_Man")
        with col3: input_min_valor = st.number_input(label="Valor mínimo", min_value=min_valor, max_value=max_valor, value=min_valor, key="Min_Valor_Man")
        with col4: input_max_valor = st.number_input(label="Valor máximo", min_value=min_valor, max_value=max_valor, value=max_valor, key="Max_Valor_Man")

        filtrar = st.button("Atualizar")
        if filtrar or 'edit_ledger' not in st.session_state:
            st.session_state.edit_ledger = filtrar_df(ledger, st.session_state["Datas_Man"], st.session_state["Min_Valor_Man"], st.session_state["Max_Valor_Man"], st.session_state["Pesquisa_Man"])
            st.session_state.edit_bank = filtrar_df(bank, st.session_state["Datas_Man"], st.session_state["Min_Valor_Man"], st.session_state["Max_Valor_Man"], st.session_state["Pesquisa_Man"])

    matrix_container = st.empty()
    matrizes_manual(matrix_container)

    nota = st.text_input("Nota")
    fazer_match = st.button("Conciliar", key=f'conciliar_bnt_{st.session_state.version}')

    if fazer_match: 
        st.session_state.version += 1
        conciliar(st.session_state.cnt, st.session_state.bnk, nota)

        for state in [f'ledger_matrix_{st.session_state.version-1}', f'bank_matrix_{st.session_state.version-1}', 'cnt', 'bnk', 'edit_ledger', 'edit_bank']:
            del st.session_state[state]
        matrix_container.empty()
        st.rerun()
    
    if st.button("Preparar ficheiro download"):
        excel_bytes = download_file()
        st.download_button(
            label="Download ficheiro",
            data=excel_bytes,
            file_name='reconciliacao.xlsx',
            mime='application/vnd.ms-excel'
        )


def reconciliacao_inicial():
        reconciliation_file = st.session_state.file
        ledger_df = reconciliation_file["CNT"]
        bank_df = reconciliation_file["BANCO"]
        ledger_df = preprocess(ledger_df)
        bank_df = preprocess(bank_df)

        conciliados, ledger, bank = match_transactions(ledger=ledger_df, bank=bank_df, date_tolerance_days=st.session_state.days)

        st.session_state.ledger = ledger
        st.session_state.bank = bank
        st.session_state.conciliados = conciliados
        st.session_state.conciliacao_inicial = 1


def remover_conciliacao(tbl):
    col_cnt = ['DataMovimento', 'DescricaoMovimento', 'ValorMovimento']
    col_bnk = ['DataExtrato', 'DescricaoBanco', 'ValorBanco']
    
    tbl_processada = pd.DataFrame(tbl)
    linhas_a_remover = tbl_processada[tbl_processada['Remover']==True]
    linhas_a_remover.drop(["Remover", "Nota"],axis=1, inplace=True)

    for col in linhas_a_remover.columns:
        linhas_a_remover[col] = linhas_a_remover[col].apply(
            lambda x: ast.literal_eval(x) if isinstance(x, str) and x.startswith("[") else x
        )

    linhas_a_remover_cnt = linhas_a_remover[col_cnt].copy()
    linhas_a_remover_banco = linhas_a_remover[col_bnk].copy()

    lista_adicionar_cnt = adicionar_linhas(linhas_a_remover_cnt)
    lista_adicionar_banco = adicionar_linhas(linhas_a_remover_banco)

    df_adicionar_cnt = pd.DataFrame(lista_adicionar_cnt)
    df_adicionar_banco = pd.DataFrame(lista_adicionar_banco)

    st.session_state.bank = pd.concat([st.session_state.bank, df_adicionar_banco], ignore_index=True)
    st.session_state.ledger = pd.concat([st.session_state.ledger, df_adicionar_cnt], ignore_index=True)

    idx_a_remover = linhas_a_remover['OrdemMovimento'].to_list()
    st.session_state.conciliados = [linha for linha in st.session_state.conciliados if linha.get("OrdemMovimento") not in idx_a_remover]
    st.rerun()


def adicionar_linhas(df):
    lista_adicionar = []
    cols = ["Data", "Descrição", "Valor"]

    for i, row in df.iterrows():
        if isinstance(row[0], list):
            num_docs = len(row[0])
            for linha in range(num_docs):
                row_dict = {
                    cols[0]: pd.to_datetime(row[0][linha]),
                    cols[1]: row[1][linha],
                    cols[2]: row[2][linha],
                    "Ver":False
                }
                lista_adicionar.append(row_dict)
        else:
            row_dict = {
                cols[0]: pd.to_datetime(row[0]),
                cols[1]: row[1],
                cols[2]: row[2],
                "Ver":False
            }
            lista_adicionar.append(row_dict)

    return lista_adicionar


def conciliados():
    lista_keys = ['OrdemMovimento', 'DataMovimento', 'DescricaoMovimento','ValorMovimento', 'DataExtrato', 'DescricaoBanco', 'ValorBanco', 'Nota']
    conciliados_show =  [
        {k: d[k] for k in lista_keys if k in d}
        for d in st.session_state.conciliados
    ]

    for row in conciliados_show:
        row['Remover'] = False

    tbl_conciliados = st.data_editor(
        data=conciliados_show
    )

    if st.button("Remover"):
        remover_conciliacao(tbl_conciliados)


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

    for data in dados:
        contagem = 0
        valor = 0
        for i, row in data.iterrows():
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
        st.metric(label="Diferença", value=f'{round((valor_por_conciliar[0]+st.session_state.saldo_i_cnt)-(st.session_state.saldo_i_banco+valor_por_conciliar[1]),2)} €')


def app():
    st.set_page_config(page_title="Reconciliações", page_icon=':star', layout='wide')
    placeholder = st.empty()

    if "file" not in st.session_state:
        st.session_state["file"] = None
        st.session_state["conciliacao_inicial"] = None

    if st.session_state["file"] is None:
        upload_file(placeholder)
        if st.session_state.file is not None:
            placeholder.empty()
            st.rerun()

    if st.session_state.file is not None and st.session_state.conciliacao_inicial == None:
        reconciliacao_inicial()

    if st.session_state.conciliacao_inicial == 1:
        st.session_state.ledger['Ver'] = False
        st.session_state.bank['Ver'] = False

        opcao = st.selectbox("", options=['Resumo', 'Conciliados', 'Manual', 'Banco - Ver opções', 'Contabilidade - Ver opções'])

        if opcao == 'Conciliados':
            conciliados()
        elif opcao == 'Manual':
            manual(st.session_state.ledger, st.session_state.bank)
        elif opcao == 'Banco - Ver opções':
            ver_opcoes_reconciliacao(st.session_state.bank)
        elif opcao == 'Contabilidade - Ver opções':
            ver_opcoes_reconciliacao(st.session_state.ledger)
        elif opcao == "Resumo":
            resumo()


app()