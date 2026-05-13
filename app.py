import streamlit as st
import pandas as pd
import sqlite3
from datetime import date

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="APAE | Prestação de Contas",
    layout="wide",
    page_icon="📄"
)

st.markdown("## 📄 Prestação de Contas - APAE Sidrolândia/MS")
st.caption("Processo: 004/2026 | Parceria: 111/2025 | CNPJ: 33.153.156/0001-61")
st.divider()

# ==========================================
# 2. BANCO DE DADOS (SQLite)
# ==========================================
DB_PATH = "apae_contas.db"

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        # Tabela para o Anexo 3 (Pagamentos)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS pagamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parcela TEXT NOT NULL,
                favorecido TEXT NOT NULL,
                endereco TEXT,
                cpf_cnpj TEXT NOT NULL,
                tipo_doc TEXT NOT NULL,
                num_doc TEXT,
                data_pagamento TEXT NOT NULL,
                valor REAL NOT NULL
            )
        """)
        # Tabela para as Receitas (Anexo 2 e 5)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS receitas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parcela TEXT NOT NULL,
                descricao TEXT NOT NULL,
                data_recebimento TEXT,
                valor REAL NOT NULL
            )
        """)
init_db()

# ==========================================
# 3. INTERFACE PRINCIPAL (Abas)
# ==========================================
tab_receitas, tab_pagamentos, tab_relatorios = st.tabs([
    "💰 Receitas (Anexo 2)", 
    "🧾 Pagamentos (Anexo 3)", 
    "🖨️ Gerar PDF Oficial"
])

# -- ABA 1: RECEITAS --
with tab_receitas:
    st.markdown("#### Lançamento de Receitas")
    with st.form("form_receita", clear_on_submit=False):
        c1, c2 = st.columns(2)
        parcela_rec = c1.text_input("Nº da Parcela (Ex: 02/10)")
        desc_rec = c2.selectbox("Origem do Recurso", ["Valor Recebido (Prefeitura)", "Rendimentos de Aplicação", "Contrapartida"])
        
        c3, c4 = st.columns(2)
        data_rec = c3.date_input("Data do Crédito")
        valor_rec = c4.number_input("Valor (R$)", min_value=0.0, step=100.0)
        
        if st.form_submit_button("Registrar Receita", type="primary"):
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("INSERT INTO receitas (parcela, descricao, data_recebimento, valor) VALUES (?,?,?,?)",
                             (parcela_rec, desc_rec, str(data_rec), valor_rec))
            st.success("Receita registrada!")

# -- ABA 2: PAGAMENTOS (ANEXO 3) --
with tab_pagamentos:
    st.markdown("#### Relação de Pagamentos (Anexo 3)")
    with st.form("form_pagamento", clear_on_submit=False):
        parcela_pag = st.text_input("Referente à Parcela (Ex: 02/10)", key="parc_pag")
        
        c1, c2 = st.columns([2, 1])
        favorecido = c1.text_input("Nome do Favorecido (Ex: ELIANE DOS SANTOS CARVALHO)")
        cpf_cnpj = c2.text_input("CPF / CNPJ")
        
        endereco = st.text_input("Endereço Completo")
        
        c3, c4, c5 = st.columns(3)
        tipo_doc = c3.selectbox("Tipo de Documento", ["Holerite", "Nota Fiscal", "RPA", "Outros"])
        num_doc = c4.text_input("Nº Doc / Mês Ref. (Ex: 02/2026)")
        data_pag = c5.date_input("Data do Pagamento", key="data_pag")
        
        valor_pag = st.number_input("Valor Pago (R$)", min_value=0.0, step=100.0, key="val_pag")
        
        if st.form_submit_button("Registrar Pagamento", type="primary"):
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("""
                    INSERT INTO pagamentos (parcela, favorecido, endereco, cpf_cnpj, tipo_doc, num_doc, data_pagamento, valor)
                    VALUES (?,?,?,?,?,?,?,?)
                """, (parcela_pag, favorecido, endereco, cpf_cnpj, tipo_doc, num_doc, str(data_pag), valor_pag))
            st.success("Pagamento registrado!")

# -- ABA 3: RELATÓRIOS --
with tab_relatorios:
   # ==========================================
# FUNÇÕES DE APOIO PARA O PDF
# ==========================================
from io import BytesIO
from xhtml2pdf import pisa

def formatar_br(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def gerar_pdf(html_content: str):
    result = BytesIO()
    pdf = pisa.pisaDocument(BytesIO(html_content.encode("UTF-8")), result)
    return result.getvalue() if not pdf.err else None

# ==========================================
# CONTINUAÇÃO DA ABA 3: RELATÓRIOS
# ==========================================
with tab_relatorios:
    st.markdown("#### 🖨️ Gerar Prestação de Contas (Anexos 1, 2, 3 e 5)")
    
    # 1. Selecionar qual parcela vamos gerar
    with sqlite3.connect(DB_PATH) as conn:
        df_pag = pd.read_sql("SELECT * FROM pagamentos", conn)
        df_rec = pd.read_sql("SELECT * FROM receitas", conn)
    
    parcelas_disponiveis = df_pag['parcela'].unique().tolist()
    
    if not parcelas_disponiveis:
        st.warning("Nenhum pagamento registrado ainda. Vá para a aba 'Pagamentos' e registre as despesas.")
    else:
        parcela_selecionada = st.selectbox("Selecione a Parcela para gerar o PDF:", parcelas_disponiveis)
        
        if st.button("Gerar PDF Unificado", type="primary", width="stretch"):
            # Filtrar dados da parcela selecionada
            pag_parcela = df_pag[df_pag['parcela'] == parcela_selecionada]
            rec_parcela = df_rec[df_rec['parcela'] == parcela_selecionada]
            
            # Cálculos Financeiros
            total_despesas = pag_parcela['valor'].sum()
            total_receitas = rec_parcela['valor'].sum()
            saldo = total_receitas - total_despesas
            
            # Montar as linhas da tabela do Anexo 3 (Pagamentos)
            linhas_anexo3 = ""
            for i, row in pag_parcela.iterrows():
                linhas_anexo3 += f"""
                <tr>
                    <td align="center">{str(i+1).zfill(2)}</td>
                    <td>{row['favorecido']}<br/><small>{row['endereco']}</small></td>
                    <td>{row['cpf_cnpj']}</td>
                    <td>{row['tipo_doc']}</td>
                    <td>{row['num_doc']}</td>
                    <td>{row['data_pagamento']}</td>
                    <td align="right">R$ {formatar_br(row['valor'])}</td>
                </tr>
                """

            # ==========================================
            # ESTRUTURA HTML DO PDF
            # ==========================================
            html_pdf = f"""
            <html>
            <head>
                <style>
                    @page {{ size: a4 portrait; margin: 1.5cm; }}
                    body {{ font-family: Helvetica, sans-serif; font-size: 10px; color: #000; }}
                    .cabecalho {{ text-align: center; font-weight: bold; margin-bottom: 20px; font-size: 12px; }}
                    .titulo-anexo {{ text-align: center; font-weight: bold; font-size: 14px; background-color: #f0f0f0; padding: 5px; border: 1px solid #000; margin-bottom: 15px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
                    th, td {{ border: 1px solid #000; padding: 5px; vertical-align: top; }}
                    th {{ background-color: #f9f9f9; text-align: left; }}
                    .assinaturas {{ width: 100%; margin-top: 50px; text-align: center; border: none; }}
                    .assinaturas td {{ border: none; padding: 20px 10px 0 10px; }}
                    .linha-ass {{ border-top: 1px solid #000; padding-top: 5px; font-weight: bold; }}
                </style>
            </head>
            <body>

                <div class="cabecalho">
                    ESTADO DE MATO GROSSO DO SUL<br/>
                    PREFEITURA DE SIDROLÂNDIA<br/>
                    Controladoria Geral do Município
                </div>
                <div class="titulo-anexo">PRESTAÇÃO DE CONTAS PARCIAL<br/>RELATÓRIO DE EXECUÇÃO FÍSICA - ANEXO 1</div>
                
                <table>
                    <tr><td colspan="2"><b>1 - NOME DO ÓRGÃO OU ENTIDADE:</b> APAE - ASSOCIAÇÃO DE PAIS E AMIGOS DOS EXCEPCIONAIS DE SIDROLANDIA MS</td><td><b>2 - UF:</b> MS</td></tr>
                    <tr><td colspan="2"><b>3 - CNPJ:</b> 33.153.156/0001-61</td><td><b>4 - N° DO PROCESSO:</b> 004/2026</td></tr>
                    <tr><td><b>5 - N° PARCERIA:</b> 111/2025</td><td><b>6 - PARCELA:</b> {parcela_selecionada}</td><td><b>7 - EXERCÍCIO:</b> 2026</td></tr>
                </table>
                <table>
                    <tr><th>AÇÃO / ESPECIFICAÇÃO</th><th>UNIDADE</th><th>QUANTIDADE EXECUTADA</th></tr>
                    <tr>
                        <td>Capacitar, integrar e apoiar a pessoa com deficiência e sua família, visando uma melhor compreensão de seus papéis...</td>
                        <td align="center">UND</td>
                        <td align="center">1.328</td>
                    </tr>
                </table>
                <table class="assinaturas">
                    <tr>
                        <td><div class="linha-ass">TULIO AUGUSTO GIMELLI<br/>CONTADOR - CRC/MS-4051</div></td>
                        <td><div class="linha-ass">SANDRA IONE STRALIOTTO SPOHR<br/>CPF: 322.032.171-20</div></td>
                        <td><div class="linha-ass">FLAVIO RODRIGUES DE SOUSA<br/>PRESIDENTE - CPF 761.909.606-00</div></td>
                    </tr>
                </table>

                <pdf:nextpage />

                <div class="cabecalho">
                    ESTADO DE MATO GROSSO DO SUL<br/>PREFEITURA DE SIDROLÂNDIA
                </div>
                <div class="titulo-anexo">PRESTAÇÃO FINANCEIRA (RECEITA E DESPESA) - ANEXO 2</div>
                
                <table>
                    <tr><td colspan="2"><b>1 - ENTIDADE:</b> APAE SIDROLÂNDIA MS</td><td><b>CNPJ:</b> 33.153.156/0001-61</td></tr>
                    <tr><td><b>PARCERIA:</b> 111/2025</td><td><b>PARCELA:</b> {parcela_selecionada}</td><td><b>TOTAL RECEITAS:</b> R$ {formatar_br(total_receitas)}</td></tr>
                </table>
                <table>
                    <tr>
                        <th>ESPECIFICAÇÃO</th>
                        <th>RECEITA EFETIVADA</th>
                        <th>DESPESA REALIZADA</th>
                        <th>SALDO</th>
                    </tr>
                    <tr>
                        <td>Execução da Parceria (Apoio à pessoa com deficiência)</td>
                        <td>R$ {formatar_br(total_receitas)}</td>
                        <td>R$ {formatar_br(total_despesas)}</td>
                        <td>R$ {formatar_br(saldo)}</td>
                    </tr>
                </table>
                <table class="assinaturas">
                    <tr>
                        <td><div class="linha-ass">TULIO AUGUSTO GIMELLI<br/>CONTADOR</div></td>
                        <td><div class="linha-ass">FLAVIO RODRIGUES DE SOUSA<br/>PRESIDENTE</div></td>
                    </tr>
                </table>

                <pdf:nextpage />

                <div class="cabecalho">
                    ESTADO DE MATO GROSSO DO SUL<br/>PREFEITURA DE SIDROLÂNDIA
                </div>
                <div class="titulo-anexo">RELAÇÃO DE PAGAMENTOS EFETUADOS - ANEXO 3</div>
                
                <table>
                    <tr>
                        <th>N°</th>
                        <th>FAVORECIDO / ENDEREÇO</th>
                        <th>CPF/CNPJ</th>
                        <th>DOCUMENTO</th>
                        <th>N° DOC</th>
                        <th>DATA PAG.</th>
                        <th>VALOR (R$)</th>
                    </tr>
                    {linhas_anexo3}
                    <tr>
                        <td colspan="6" align="right"><b>TOTAL GERAL DA PARCELA:</b></td>
                        <td align="right"><b>R$ {formatar_br(total_despesas)}</b></td>
                    </tr>
                </table>

                <pdf:nextpage />

                <div class="cabecalho">
                    ESTADO DE MATO GROSSO DO SUL<br/>PREFEITURA DE SIDROLÂNDIA
                </div>
                <div class="titulo-anexo">CONCILIAÇÃO BANCÁRIA - ANEXO 5</div>
                
                <table>
                    <tr><td><b>CONVENENTE:</b> APAE - ASSOCIAÇÃO DE PAIS E AMIGOS DOS EXCEPCIONAIS</td><td><b>PARCERIA N°:</b> 111/2025</td></tr>
                </table>
                <table>
                    <tr>
                        <th>DESCRIÇÃO</th>
                        <th>VALOR (R$)</th>
                    </tr>
                    <tr>
                        <td>1. Total de Transferências / Receitas da Parcela</td>
