import calendar
import datetime
import io
import json
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
import pandas as pd
import streamlit as st
import holidays
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

# Tenta importar streamlit_calendar com graceful fallback
try:
    from streamlit_calendar import calendar as st_calendar
    HAS_CALENDAR_COMPONENT = True
except ImportError:
    HAS_CALENDAR_COMPONENT = False

# Configuração da página
st.set_page_config(
    page_title="DETRAN/MA - Gestão de Exames Práticos",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilização CSS Customizada
st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; }
    .header-container {
        background: linear-gradient(90deg, #1A365D 0%, #2B6CB0 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header-container h1 { color: white !important; margin: 0; font-size: 26px; font-weight: 700; }
    .header-container p { color: #E2E8F0; margin: 5px 0 0 0; font-size: 14px; }
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border-left: 5px solid #2B6CB0;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        white-space: pre-wrap;
        background-color: #EDF2F7;
        border-radius: 6px 6px 0 0;
        color: #2D3748;
        font-weight: 600;
        padding: 10px 16px;
    }
    .stTabs [aria-selected="true"] { background-color: #1A365D !important; color: white !important; }
    .stButton>button { border-radius: 6px; font-weight: 600; }
    </style>
""",
    unsafe_allow_html=True,
)

# Lista completa de meses
MESES_LISTA = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
]
ANOS_LISTA = [2026, 2027, 2028, 2029, 2030]
DIAS_SEMANA_OPCOES = [
    "Segunda",
    "Terça",
    "Quarta",
    "Quinta",
    "Sexta",
    "Sábado",
    "Domingo",
]

# --- ESTADO DA SESSÃO ---
if "bancas_config" not in st.session_state:
    st.session_state["bancas_config"] = {
        "Banca São Luís": [
            "São Luís Pátio",
            "São Luís Cohatrac",
            "São Luís Castelinho",
            "São Luís Cidade Operária",
            "Paço do Lumiar",
            "Raposa",
            "São José de Ribamar",
            "Pinheiro",
            "São Bento",
            "Carutapera",
            "Turilândia",
            "Tutóia",
            "Barreirinhas",
            "Chapadinha",
            "Axixá",
            "Rosário",
            "Santa Rita",
            "Coroatá",
            "Itapecuru-mirim",
            "Codó",
        ],
        "Banca Imperatriz": [
            "Imperatriz",
            "Açailândia",
            "Estreito",
            "Grajaú",
            "Amarante do Maranhão",
        ],
        "Banca Timon": ["Timon - Pátio", "Matões", "Parnarama"],
        "Banca Caxias": [
            "Caxias - Pátio",
            "Codó - Regional",
            "São João do Sóter",
            "Aldeias Altas",
        ],
        "Banca Bacabal": [
            "Bacabal - Pátio",
            "Pedreiras",
            "Lago da Pedra",
            "Olho d'Água das Cunhãs",
        ],
        "Banca Santa Inês": [
            "Santa Inês - Pátio",
            "Viana",
            "Zé Doca",
            "Monção",
            "Pindaré-Mirim",
        ],
    }

if "lista_horarios" not in st.session_state:
    st.session_state["lista_horarios"] = [
        "08:30",
        "09:30",
        "10:30",
        "11:30",
        "13:30",
        "14:30",
        "15:30",
        "16:30",
    ]

if "historico_localidades" not in st.session_state:
    st.session_state["historico_localidades"] = {}

if "feriados_locais_dict" not in st.session_state:
    st.session_state["feriados_locais_dict"] = {}

if "viagens_registradas" not in st.session_state:
    st.session_state["viagens_registradas"] = []

if "dias_permitidos_dict" not in st.session_state:
    st.session_state["dias_permitidos_dict"] = {}

CAPACIDADE_BANCAS_PADRAO = {
    "Banca São Luís": 34,
    "Banca Imperatriz": 14,
    "Banca Bacabal": 5,
    "Banca Caxias": 4,
    "Banca Timon": 4,
    "Banca Santa Inês": 3,
}
LIMITE_FORA_SEDE_PADRAO = {
    "Banca São Luís": 16,
    "Banca Imperatriz": 10,
    "Banca Bacabal": 5,
    "Banca Caxias": 4,
    "Banca Timon": 4,
    "Banca Santa Inês": 3,
}
if "capacidade_bancas" not in st.session_state:
    st.session_state["capacidade_bancas"] = dict(CAPACIDADE_BANCAS_PADRAO)
if "limite_fora_sede_bancas" not in st.session_state:
    st.session_state["limite_fora_sede_bancas"] = dict(LIMITE_FORA_SEDE_PADRAO)
if "disponibilidade_semanal" not in st.session_state:
    st.session_state["disponibilidade_semanal"] = {}
if "controle_capacidade_ativo" not in st.session_state:
    st.session_state["controle_capacidade_ativo"] = False

# --- BACKUP AUTOMÁTICO EM DISCO ---
ARQUIVO_AUTOSAVE = "detran_autosave_backup.json"


def _numero_seguro_para_json(valor):
    """Converte tipos numéricos do NumPy/Pandas para tipos nativos do Python."""
    if hasattr(valor, "item"):
        return valor.item()
    return str(valor)


def montar_estado_para_backup():
    """Reúne tudo que precisa ser preservado num dicionário simples para JSON."""
    viagens_serializadas = []
    for v in st.session_state["viagens_registradas"]:
        v_copia = dict(v)
        v_copia["Data Inicio"] = v["Data Inicio"].isoformat()
        v_copia["Data Fim"] = v["Data Fim"].isoformat()
        viagens_serializadas.append(v_copia)

    historico_serializado = {
        chave: df.to_dict(orient="records")
        for chave, df in st.session_state["historico_localidades"].items()
    }

    return {
        "bancas_config": st.session_state["bancas_config"],
        "lista_horarios": st.session_state["lista_horarios"],
        "horarios_inativos": st.session_state.get("horarios_inativos", []),
        "historico_localidades": historico_serializado,
        "feriados_locais_dict": st.session_state["feriados_locais_dict"],
        "viagens_registradas": viagens_serializadas,
        "dias_permitidos_dict": st.session_state["dias_permitidos_dict"],
        "capacidade_bancas": st.session_state["capacidade_bancas"],
        "limite_fora_sede_bancas": st.session_state["limite_fora_sede_bancas"],
        "disponibilidade_semanal": st.session_state["disponibilidade_semanal"],
        "controle_capacidade_ativo": st.session_state.get("controle_capacidade_ativo", False),
        "schema_version": 3,
    }


def aplicar_estado_do_backup(dados):
    """Repõe o conteúdo de um backup no session_state atual."""
    st.session_state["bancas_config"] = dados.get(
        "bancas_config", st.session_state["bancas_config"]
    )
    st.session_state["lista_horarios"] = dados.get(
        "lista_horarios", st.session_state["lista_horarios"]
    )
    st.session_state["horarios_inativos"] = dados.get("horarios_inativos", [])
    st.session_state["historico_localidades"] = {
        chave: pd.DataFrame(registros)
        for chave, registros in dados.get("historico_localidades", {}).items()
    }
    st.session_state["feriados_locais_dict"] = dados.get("feriados_locais_dict", {})

    viagens_restauradas = []
    for v in dados.get("viagens_registradas", []):
        v_copia = dict(v)
        v_copia["Data Inicio"] = datetime.date.fromisoformat(v["Data Inicio"])
        v_copia["Data Fim"] = datetime.date.fromisoformat(v["Data Fim"])
        viagens_restauradas.append(v_copia)
    st.session_state["viagens_registradas"] = viagens_restauradas

    st.session_state["dias_permitidos_dict"] = dados.get("dias_permitidos_dict", {})
    st.session_state["capacidade_bancas"] = {**CAPACIDADE_BANCAS_PADRAO, **dados.get("capacidade_bancas", {})}
    st.session_state["limite_fora_sede_bancas"] = {**LIMITE_FORA_SEDE_PADRAO, **dados.get("limite_fora_sede_bancas", {})}
    st.session_state["disponibilidade_semanal"] = dados.get("disponibilidade_semanal", {})
    st.session_state["controle_capacidade_ativo"] = bool(dados.get("controle_capacidade_ativo", False))


def salvar_autosave_em_disco():
    """Salva o estado atual em um arquivo JSON local."""
    try:
        tmp = ARQUIVO_AUTOSAVE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(montar_estado_para_backup(), f, ensure_ascii=False, default=_numero_seguro_para_json)
        os.replace(tmp, ARQUIVO_AUTOSAVE)
        st.session_state["ultimo_autosave"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        st.session_state["erro_autosave"] = None
    except Exception as e:
        st.session_state["erro_autosave"] = str(e)


def carregar_autosave_do_disco():
    """Carrega o arquivo de backup automático se existir."""
    if not os.path.exists(ARQUIVO_AUTOSAVE):
        return False
    try:
        with open(ARQUIVO_AUTOSAVE, "r", encoding="utf-8") as f:
            dados = json.load(f)
        aplicar_estado_do_backup(dados)
        return True
    except Exception:
        return False


if "autosave_carregado_nesta_sessao" not in st.session_state:
    carregar_autosave_do_disco()
    st.session_state["autosave_carregado_nesta_sessao"] = True

# Garante numeração para viagens que não possuem o atributo de banca itinerante
contador_numero_por_banca = {}
for v in st.session_state["viagens_registradas"]:
    if "Numero Banca Itinerante" not in v or not v["Numero Banca Itinerante"]:
        contador_numero_por_banca[v["Banca"]] = (
            contador_numero_por_banca.get(v["Banca"], 0) + 1
        )
        v["Numero Banca Itinerante"] = contador_numero_por_banca[v["Banca"]]
    else:
        contador_numero_por_banca[v["Banca"]] = max(
            contador_numero_por_banca.get(v["Banca"], 0),
            v["Numero Banca Itinerante"],
        )

# Migração de dados de Imperatriz
NOME_ANTIGO_IMPERATRIZ = "Imperatriz - Pátio"
NOME_NOVO_IMPERATRIZ = "Imperatriz"
if NOME_ANTIGO_IMPERATRIZ in st.session_state["bancas_config"].get(
    "Banca Imperatriz", []
):
    st.session_state["bancas_config"]["Banca Imperatriz"] = [
        NOME_NOVO_IMPERATRIZ if loc == NOME_ANTIGO_IMPERATRIZ else loc
        for loc in st.session_state["bancas_config"]["Banca Imperatriz"]
    ]
    st.session_state["historico_localidades"] = {
        chave.replace(NOME_ANTIGO_IMPERATRIZ, NOME_NOVO_IMPERATRIZ): df
        for chave, df in st.session_state["historico_localidades"].items()
    }
    for v in st.session_state["viagens_registradas"]:
        if v["Destino"] == NOME_ANTIGO_IMPERATRIZ:
            v["Destino"] = NOME_NOVO_IMPERATRIZ
    for chave_dias in list(st.session_state["dias_permitidos_dict"].keys()):
        if chave_dias.endswith(f"_{NOME_ANTIGO_IMPERATRIZ}"):
            nova_chave = chave_dias.replace(
                NOME_ANTIGO_IMPERATRIZ, NOME_NOVO_IMPERATRIZ
            )
            st.session_state["dias_permitidos_dict"][
                nova_chave
            ] = st.session_state["dias_permitidos_dict"].pop(chave_dias)
    for chave_feriado in list(st.session_state["feriados_locais_dict"].keys()):
        if chave_feriado.endswith(f"_{NOME_ANTIGO_IMPERATRIZ}"):
            nova_chave = chave_feriado.replace(
                NOME_ANTIGO_IMPERATRIZ, NOME_NOVO_IMPERATRIZ
            )
            st.session_state["feriados_locais_dict"][
                nova_chave
            ] = st.session_state["feriados_locais_dict"].pop(chave_feriado)


def obter_sede(banca):
    """Retorna a sede da banca (primeiro item cadastrado)."""
    lista = st.session_state["bancas_config"].get(banca, [])
    return lista[0] if lista else None


def periodo_do_horario(horario):
    """Classifica horário em Manhã ou Tarde."""
    try:
        h, m = map(int, str(horario).split(":"))
        return "Manhã" if (h, m) < (12, 0) else "Tarde"
    except Exception:
        return "Manhã"


def viagem_cobre_periodo(v, data, periodo):
    if not (v["Data Inicio"] <= data <= v["Data Fim"]):
        return False
    turno = v.get("Turno", "Dia inteiro")
    return turno == "Dia inteiro" or turno == periodo


def banca_vinculada(v, banca):
    return v.get("Banca") == banca or (v.get("Apoio Conjunto") and banca in ["Banca Caxias", "Banca Timon"])


def examinadores_da_banca_na_viagem(v, banca):
    if v.get("Apoio Conjunto"):
        if banca == "Banca Caxias":
            return int(v.get("Examinadores Caxias", v.get("Examinadores", 0)))
        if banca == "Banca Timon":
            return int(v.get("Examinadores Timon", v.get("Examinadores", 0)))
    return int(v.get("Examinadores", 0)) if v.get("Banca") == banca else 0


def viagens_ativas_no_periodo(banca, data, periodo, ignorar_idx=None):
    resultado = []
    for idx, v in enumerate(st.session_state["viagens_registradas"]):
        if ignorar_idx is not None and idx == ignorar_idx:
            continue
        if banca_vinculada(v, banca) and viagem_cobre_periodo(v, data, periodo):
            resultado.append(v)
    return resultado


def semana_chave(data):
    segunda = data - datetime.timedelta(days=data.weekday())
    return segunda.isoformat()


def disponibilidade_semanal_banca(banca, data):
    chave = f"{banca}|{semana_chave(data)}"
    return int(st.session_state["disponibilidade_semanal"].get(chave, st.session_state["capacidade_bancas"].get(banca, 0)))


def conflitos_de_horario_rota(banca, destino, inicio, fim, turno, numero, apoio=False, ignorar_idx=None):
    conflitos = []
    for idx, v in enumerate(st.session_state["viagens_registradas"]):
        if ignorar_idx is not None and idx == ignorar_idx:
            continue
        if v.get("Numero Banca Itinerante") != numero:
            continue
        if not banca_vinculada(v, banca):
            continue
        ini = max(inicio, v["Data Inicio"]); f = min(fim, v["Data Fim"])
        if ini > f:
            continue
        t2 = v.get("Turno", "Dia inteiro")
        if turno == "Dia inteiro" or t2 == "Dia inteiro" or turno == t2:
            conflitos.append(v)
    return conflitos


def extrair_data_feriado(linha, ano_ref):
    if not linha or not linha.strip():
        return None
    parte_data = linha.split("-")[0].split("–")[0].split("(")[0].strip()
    formatos = ["%d/%m/%Y", "%d/%m", "%Y-%m-%d"]
    for fmt in formatos:
        try:
            dt = datetime.datetime.strptime(parte_data, fmt).date()
            if fmt == "%d/%m":
                dt = dt.replace(year=ano_ref)
            return dt
        except ValueError:
            continue
    return None


def formatar_datas_exames(dias_unicos):
    """Converte lista de dias do mês no formato de intervalos para publicação oficial."""
    if not dias_unicos:
        return "-"

    dias_ordenados = sorted(set(dias_unicos))
    grupos = []
    grupo_atual = [dias_ordenados[0]]
    for d in dias_ordenados[1:]:
        if d == grupo_atual[-1] + 1:
            grupo_atual.append(d)
        else:
            grupos.append(grupo_atual)
            grupo_atual = [d]
    grupos.append(grupo_atual)

    partes = []
    for g in grupos:
        if len(g) == 1:
            partes.append(f"{g[0]:02d}")
        elif len(g) == 2:
            partes.append(f"{g[0]:02d} e {g[1]:02d}")
        else:
            partes.append(f"{g[0]:02d} a {g[-1]:02d}")

    return "; ".join(partes)


def gerar_excel_modelo_oficial(df_resumo_completo, mes_sel, ano_sel):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Plan1"

    font_padrao = Font(name="Arial", size=10)
    align_center = Alignment(horizontal="center", vertical="center")
    align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    larguras_colunas = {
        "A": 3,
        "B": 30.29,
        "C": 8.29,
        "D": 8.0,
        "E": 6.86,
        "F": 8.43,
        "G": 8.43,
        "H": 25.29,
    }
    for col_letter, largura in larguras_colunas.items():
        ws.column_dimensions[col_letter].width = largura

    cols_categorias = ["Cat A", "Cat B", "Cat C", "Cat D", "Cat E"]
    curr_row = 3

    for banca, grp in df_resumo_completo.groupby("Banca", sort=False):
        banca_nome_pub = banca.replace("Banca ", "")

        ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=8)
        cell_titulo = ws.cell(
            row=curr_row,
            column=2,
            value=(
                f"Quantidade de Exames Mensais de {banca_nome_pub} e Região"
                f" - {mes_sel.upper()} DE {ano_sel}"
            ),
        )
        cell_titulo.font = font_padrao
        cell_titulo.alignment = align_center
        for c in range(2, 9):
            ws.cell(row=curr_row, column=c).border = border_all
        curr_row += 1

        headers = ["Cidade"] + cols_categorias + ["Datas dos exames"]
        for idx_col, h_text in enumerate(headers, 2):
            cell = ws.cell(row=curr_row, column=idx_col, value=h_text)
            cell.font = font_padrao
            cell.alignment = align_center
            cell.border = border_all
        curr_row += 1

        for _, r in grp.iterrows():
            cell_cidade = ws.cell(row=curr_row, column=2, value=str(r["Localidade"]))
            cell_cidade.font = font_padrao
            cell_cidade.alignment = align_left
            cell_cidade.border = border_all

            for idx_col, cat_col in enumerate(cols_categorias, 3):
                valor = int(r.get(cat_col, 0) or 0)
                cell = ws.cell(
                    row=curr_row, column=idx_col, value=valor if valor > 0 else None
                )
                cell.font = font_padrao
                cell.alignment = align_center
                cell.number_format = "#,##0"
                cell.border = border_all

            cell_datas = ws.cell(
                row=curr_row, column=8, value=str(r.get("Datas dos exames", "-"))
            )
            cell_datas.font = font_padrao
            cell_datas.alignment = align_center_wrap
            cell_datas.number_format = "@"
            cell_datas.border = border_all

            curr_row += 1

        curr_row += 1

    ultima_linha_com_dados = curr_row - 2
    ws.print_area = f"A3:H{max(ultima_linha_com_dados, 3)}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 0.51
    ws.page_margins.right = 0.51
    ws.page_margins.top = 0.79
    ws.page_margins.bottom = 0.79
    ws.page_margins.header = 0.31
    ws.page_margins.footer = 0.31

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_pdf_oficial_enxuto(df_dados, banca, local, mes, ano):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15,
    )
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Heading1"],
        fontSize=12,
        alignment=1,
        spaceAfter=6,
    )
    sub_style = ParagraphStyle(
        "SubStyle", parent=styles["Normal"], fontSize=9, alignment=1, spaceAfter=12
    )

    elements.append(
        Paragraph(
            "<b>DETRAN-MA — DEPARTAMENTO ESTADUAL DE TRÂNSITO DO MARANHÃO</b>",
            title_style,
        )
    )
    elements.append(
        Paragraph(
            f"GRADE DE LANÇAMENTO DE VAGAS DE EXAMES PRÁTICOS — {banca.upper()} ({local.upper()}) — {mes}/{ano}",
            sub_style,
        )
    )

    df_filtrado = df_dados[
        (df_dados["Status"] == "Disponível") & (df_dados["Total"] > 0)
    ].copy()

    pcd_cols = ["PCD A", "PCD B", "PCD C", "PCD D", "PCD E"]
    pcd_usadas = (
        [
            col
            for col in pcd_cols
            if col in df_filtrado.columns and df_filtrado[col].sum() > 0
        ]
        if not df_filtrado.empty
        else []
    )

    cols_vagas_todas = ["Cat A", "Cat B", "Cat C", "Cat D", "Cat E"] + pcd_usadas
    cols_exibir = (
        [
            "Data",
            "Dia da Semana",
            "Status",
            "Exam. M",
            "Exam. T",
            "Horário",
            "Cat A",
            "Cat B",
            "Cat C",
            "Cat D",
            "Cat E",
        ]
        + pcd_usadas
        + ["Total"]
    )

    if df_filtrado.empty:
        elements.append(
            Paragraph(
                "<b>Nenhuma vaga ofertada para os parâmetros selecionados.</b>",
                sub_style,
            )
        )
    else:
        rows_com_totais = []
        indices_totais = []

        for data_val, group in df_filtrado.groupby("Data", sort=False):
            for _, r in group.iterrows():
                row_list = []
                for c in cols_exibir:
                    val = r[c]
                    if c in cols_vagas_todas + ["Total"]:
                        row_list.append("-" if val == 0 else str(int(val)))
                    else:
                        row_list.append(str(val))
                rows_com_totais.append(row_list)

            tot_row = [f"TOTAL {data_val}", "", "SUBTOTAL", "-", "-", "-"]
            for c in cols_vagas_todas:
                soma_c = group[c].sum()
                tot_row.append("-" if soma_c == 0 else str(int(soma_c)))
            tot_row.append(str(int(group["Total"].sum())))

            rows_com_totais.append(tot_row)
            indices_totais.append(len(rows_com_totais))

        data_table = [cols_exibir] + rows_com_totais
        t = Table(data_table, repeatRows=1)

        table_styles = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
            ("TOPPADDING", (0, 0), (-1, 0), 4),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ]

        for idx in indices_totais:
            table_styles.append(
                ("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#EDF2F7"))
            )
            table_styles.append(
                ("FONTNAME", (0, idx), (-1, idx), "Helvetica-Bold")
            )
            table_styles.append(
                ("TEXTCOLOR", (0, idx), (-1, idx), colors.HexColor("#1A365D"))
            )

        t.setStyle(TableStyle(table_styles))
        elements.append(t)

    doc.build(elements)
    buffer.seek(0)
    return buffer


# --- HEADER ---
st.markdown(
    """
    <div class="header-container">
        <h1>🚗 DETRAN/MA — Sistema de Gestão de Exames Práticos</h1>
        <p>Planejamento de Vagas, Controle de Efetivo de Examinadores e Distribuição por Localidades</p>
    </div>
""",
    unsafe_allow_html=True,
)

# --- BARRA LATERAL: BACKUP MANUAL ---
with st.sidebar.expander("💾 Backup de Segurança", expanded=False):
    st.caption(
        "Seu progresso é salvo automaticamente a cada ação. Use os botões"
        " abaixo só se quiser guardar uma cópia extra ou restaurar um"
        " backup específico."
    )

    nome_arquivo_backup = (
        f"backup_detran_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.json"
    )
    st.download_button(
        label="⬇️ Baixar backup agora",
        data=json.dumps(
            montar_estado_para_backup(),
            ensure_ascii=False,
            default=_numero_seguro_para_json,
        ),
        file_name=nome_arquivo_backup,
        mime="application/json",
    )

    backup_enviado = st.file_uploader(
        "Restaurar um backup (.json):", type=["json"], key="upload_backup"
    )
    if backup_enviado is not None:
        if st.button("♻️ Restaurar este backup"):
            try:
                dados_backup = json.load(backup_enviado)
                aplicar_estado_do_backup(dados_backup)
                st.success("Backup restaurado com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Não foi possível ler este arquivo de backup: {e}")

if st.session_state.get("ultimo_autosave"):
    st.caption(f"💾 Último salvamento automático: {st.session_state['ultimo_autosave']}")
if st.session_state.get("erro_autosave"):
    st.warning(f"⚠️ O salvamento automático apresentou erro: {st.session_state['erro_autosave']}")

aba_quadro, aba_cal, aba_viagens, aba_horarios, aba_relatorio, aba_gestao = (
    st.tabs([
        "🗓️ Quadro Geral de Examinadores",
        "📅 Montar Calendário Detalhado",
        "🚍 Gestão de Viagens Itinerantes",
        "⏰ Horários & Turmas",
        "📊 Dashboard Consolidado",
        "➕ Gestão de Localidades",
    ])
)

# --- ABA: QUADRO GERAL DE EXAMINADORES ---
with aba_quadro:
    st.markdown("### 🗓️ Quadro Matriz de Distribuição de Examinadores")

    BANCAS_COM_QUADRO_MATRIZ = ["Banca São Luís", "Banca Imperatriz"]

    col_q1, col_q2, col_q3 = st.columns(3)
    with col_q1:
        banca_q = st.selectbox(
            "Selecione a Banca:",
            BANCAS_COM_QUADRO_MATRIZ,
            key="q_banca",
        )
    with col_q2:
        mes_q = st.selectbox(
            "Mês de Visualização:", MESES_LISTA, index=10, key="q_mes"
        )
        mes_num_q = MESES_LISTA.index(mes_q) + 1
    with col_q3:
        ano_q = st.selectbox("Ano de Visualização:", ANOS_LISTA, index=0, key="q_ano")

    st.markdown("---")

    cal = calendar.Calendar(firstweekday=0)
    dias_mes = [
        d
        for d in cal.itermonthdates(ano_q, mes_num_q)
        if d.month == mes_num_q and d.weekday() < 5
    ]

    try:
        feriados_oficiais_q = holidays.Brazil(subdiv="MA", years=ano_q)
    except TypeError:
        feriados_oficiais_q = holidays.BR(state="MA", years=ano_q)

    semanas = {}
    for d in dias_mes:
        num_semana = d.isocalendar()[1]
        if num_semana not in semanas:
            semanas[num_semana] = []
        semanas[num_semana].append(d)

    locais_banca = st.session_state["bancas_config"][banca_q]

    for idx, (sem, dias_sem) in enumerate(semanas.items(), 1):
        ini_sem = dias_sem[0].strftime("%d")
        fim_sem = dias_sem[-1].strftime("%d")
        st.markdown(f"#### 📅 Semana {idx} ({ini_sem} a {fim_sem} de {mes_q})")

        matriz_dados = []

        for loc in locais_banca:
            linha = [loc]
            chave_loc = f"{banca_q}_{mes_q}_{ano_q}_{loc}"
            df_loc = st.session_state["historico_localidades"].get(chave_loc)

            chave_dias_loc = f"{banca_q}_{loc}"
            dias_permitidos_loc = st.session_state["dias_permitidos_dict"].get(
                chave_dias_loc,
                ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"],
            )

            for d in dias_sem:
                dia_semana_str_d = DIAS_SEMANA_OPCOES[d.weekday()]
                if dia_semana_str_d not in dias_permitidos_loc:
                    linha.append("-")
                    continue

                data_str = d.strftime("%d/%m/%Y")

                viagem_ativa = next(
                    (
                        v
                        for v in st.session_state["viagens_registradas"]
                        if (
                            v["Banca"] == banca_q
                            or (
                                v.get("Apoio Conjunto")
                                and banca_q in ["Banca Caxias", "Banca Timon"]
                            )
                        )
                        and v["Destino"] == loc
                        and v["Data Inicio"] <= d <= v["Data Fim"]
                    ),
                    None,
                )

                if viagem_ativa:
                    linha.append(f"{examinadores_da_banca_na_viagem(viagem_ativa, banca_q)} (Viagem)")
                elif df_loc is not None:
                    df_dia_todos = df_loc[df_loc["Data"] == data_str]
                    eh_feriado_local = (
                        not df_dia_todos.empty
                        and (
                            df_dia_todos["Status"] == "Feriado / Sem Atendimento"
                        ).any()
                    )
                    if eh_feriado_local or d in feriados_oficiais_q:
                        linha.append("Feriado")
                        continue

                    df_dia = df_dia_todos[df_dia_todos["Status"] == "Disponível"]
                    if not df_dia.empty:
                        qtd_ex = max(df_dia["Exam. M"].max(), df_dia["Exam. T"].max())
                        linha.append(str(int(qtd_ex)))
                    else:
                        linha.append("x")
                elif d in feriados_oficiais_q:
                    linha.append("Feriado")
                else:
                    linha.append("x")
            matriz_dados.append(linha)

        linha_total = ["TOTAL EXAMINADORES EM CAMPO"]

        for d in dias_sem:
            dia_semana_str_d = DIAS_SEMANA_OPCOES[d.weekday()]
            data_str = d.strftime("%d/%m/%Y")

            total_manha = 0
            total_tarde = 0

            for loc in locais_banca:
                chave_dias_loc = f"{banca_q}_{loc}"
                dias_permitidos_loc = st.session_state["dias_permitidos_dict"].get(
                    chave_dias_loc,
                    ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"],
                )
                if dia_semana_str_d not in dias_permitidos_loc:
                    continue

                viagem_manha = next(
                    (v for v in st.session_state["viagens_registradas"]
                     if banca_vinculada(v, banca_q)
                     and v.get("Destino") == loc
                     and viagem_cobre_periodo(v, d, "Manhã")),
                    None,
                )
                viagem_tarde = next(
                    (v for v in st.session_state["viagens_registradas"]
                     if banca_vinculada(v, banca_q)
                     and v.get("Destino") == loc
                     and viagem_cobre_periodo(v, d, "Tarde")),
                    None,
                )

                chave_loc = f"{banca_q}_{mes_q}_{ano_q}_{loc}"
                df_loc = st.session_state["historico_localidades"].get(chave_loc)

                if df_loc is not None:
                    df_dia = df_loc[
                        (df_loc["Data"] == data_str) & (df_loc["Status"] == "Disponível")
                    ]
                    if not df_dia.empty:
                        if viagem_manha is None:
                            total_manha += int(df_dia["Exam. M"].max())
                        if viagem_tarde is None:
                            total_tarde += int(df_dia["Exam. T"].max())

            equipes_manha = {}
            equipes_tarde = {}
            for v in st.session_state["viagens_registradas"]:
                if not banca_vinculada(v, banca_q):
                    continue
                equipe = int(v.get("Numero Banca Itinerante", 0))
                qtd = examinadores_da_banca_na_viagem(v, banca_q)
                if qtd <= 0:
                    continue
                if viagem_cobre_periodo(v, d, "Manhã"):
                    equipes_manha[equipe] = max(equipes_manha.get(equipe, 0), qtd)
                if viagem_cobre_periodo(v, d, "Tarde"):
                    equipes_tarde[equipe] = max(equipes_tarde.get(equipe, 0), qtd)

            total_manha += sum(equipes_manha.values())
            total_tarde += sum(equipes_tarde.values())

            soma_dia = max(total_manha, total_tarde)
            linha_total.append(str(int(soma_dia)))

        matriz_dados.append(linha_total)

        df_semana = pd.DataFrame(
            matriz_dados,
            columns=["Localidade"] + [d.strftime("%d/%m - %a") for d in dias_sem],
        )
        st.dataframe(df_semana, use_container_width=True, hide_index=True)

# --- ABA: GESTÃO DE VIAGENS ITINERANTES ---
with aba_viagens:
    st.markdown("### 🚍 Controle de Viagens e Equipes Itinerantes")
    st.info(
        "Organize aqui os deslocamentos das bancas para atendimento aos"
        " municípios do interior. As equipes em viagem serão automaticamente"
        " somadas ao controle de efetivo diário e aplicadas na montagem de"
        " calendário."
    )

    v_col1, v_col2 = st.columns([1, 2])

    with v_col1:
        st.markdown("#### ➕ Cadastrar Nova Viagem")

        banca_v = st.selectbox(
            "Banca Responsável:",
            list(st.session_state["bancas_config"].keys()),
            key="v_banca",
        )

        numeros_existentes_banca_v = sorted(
            {
                v["Numero Banca Itinerante"]
                for v in st.session_state["viagens_registradas"]
                if v["Banca"] == banca_v
            }
        )

        proximo_numero_banca_v = (
            max(numeros_existentes_banca_v) + 1 if numeros_existentes_banca_v else 1
        )

        opcoes_time_itinerante = [
            f"➕ Nova banca itinerante (Banca {proximo_numero_banca_v:02d})"
        ] + [
            f"🔁 Continuar Banca {n:02d} já cadastrada"
            for n in numeros_existentes_banca_v
        ]
        escolha_time_itinerante = st.selectbox(
            "Esta viagem pertence a:",
            opcoes_time_itinerante,
            key="v_time_itinerante",
            help=(
                "Se a mesma equipe vai atender mais de um município em"
                " sequência na mesma viagem, cadastre o primeiro trecho como"
                " \"Nova banca itinerante\" e os seguintes como"
                " \"Continuar\" essa mesma banca."
            ),
        )

        if escolha_time_itinerante.startswith("➕"):
            numero_banca_itinerante_v = proximo_numero_banca_v
            examinadores_padrao_v = int(st.session_state["limite_fora_sede_bancas"].get(banca_v, st.session_state["capacidade_bancas"].get(banca_v, 1)))
        else:
            match_banca_it = re.search(r"Banca\s+(\d+)", escolha_time_itinerante)
            if not match_banca_it:
                st.error("Não foi possível identificar o número da banca itinerante selecionada.")
                st.stop()
            numero_banca_itinerante_v = int(match_banca_it.group(1))
            trecho_existente = next(
                (
                    v
                    for v in st.session_state["viagens_registradas"]
                    if v["Banca"] == banca_v
                    and v["Numero Banca Itinerante"] == numero_banca_itinerante_v
                ),
                None,
            )
            examinadores_padrao_v = (
                trecho_existente.get("Examinadores", 1) if trecho_existente else int(st.session_state["limite_fora_sede_bancas"].get(banca_v, 1))
            )

        apoio_conjunto = False
        if banca_v in ["Banca Caxias", "Banca Timon"]:
            apoio_conjunto = st.checkbox(
                "🤝 Viagem Conjunta (Caxias + Timon)",
                help=(
                    "Marque se as equipes de Caxias e Timon farão esta viagem em"
                    " parceria."
                ),
            )

        sede_banca_v = obter_sede(banca_v)
        locais_destino = [
            l
            for l in st.session_state["bancas_config"][banca_v]
            if l != sede_banca_v
        ]
        if not locais_destino:
            locais_destino = st.session_state["bancas_config"][banca_v]

        destino_v = st.selectbox(
            "Município de Destino:", locais_destino, key="v_dest"
        )

        d_inicio = st.date_input(
            "Data de Início da Viagem:",
            datetime.date(2026, 11, 9),
            format="DD/MM/YYYY",
        )
        d_fim = st.date_input(
            "Data de Término da Viagem:",
            datetime.date(2026, 11, 13),
            format="DD/MM/YYYY",
        )
        turno_v = st.selectbox(
            "Período de atendimento:",
            ["Dia inteiro", "Manhã", "Tarde"],
            help=("Use Manhã/Tarde quando a mesma equipe atender municípios diferentes no mesmo dia."),
        )

        if apoio_conjunto:
            col_c, col_t = st.columns(2)
            with col_c:
                num_exam_caxias = st.number_input("Examinadores — Caxias:", min_value=0, max_value=4, value=min(4, examinadores_padrao_v), step=1)
            with col_t:
                num_exam_timon = st.number_input("Examinadores — Timon:", min_value=0, max_value=4, value=min(4, examinadores_padrao_v), step=1)
            num_exam_v = int(num_exam_caxias if banca_v == "Banca Caxias" else num_exam_timon)
        else:
            num_exam_v = st.number_input(
                "Nº de Examinadores Deslocados:",
                min_value=1,
                max_value=50,
                value=max(1, min(50, examinadores_padrao_v)),
            )
        obs_v = st.text_input(
            "Observações / Portaria:", placeholder="Ex: Portaria nº 123/2026"
        )

        controle_capacidade_ativo = st.checkbox(
            "Ativar controle automático de capacidade/disponibilidade",
            value=bool(st.session_state.get("controle_capacidade_ativo", False)),
            key="controle_capacidade_ativo",
            help=(
                "Desativado: permite cadastrar viagens sem bloqueio por capacidade."
            ),
        )

        with st.expander("⚙️ Capacidade e disponibilidade semanal", expanded=False):
            st.caption("Ajuste manualmente os limites operacionais.")
            cap_cols = st.columns(2)
            for idx_cap, bcap in enumerate(st.session_state["bancas_config"].keys()):
                with cap_cols[idx_cap % 2]:
                    cap = st.number_input(f"{bcap} — efetivo total", min_value=0, max_value=100, value=int(st.session_state["capacidade_bancas"].get(bcap, 0)), key=f"cap_{bcap}")
                    lim = st.number_input(f"{bcap} — máximo fora da sede", min_value=0, max_value=100, value=int(st.session_state["limite_fora_sede_bancas"].get(bcap, 0)), key=f"lim_{bcap}")
                    st.session_state["capacidade_bancas"][bcap] = int(cap)
                    st.session_state["limite_fora_sede_bancas"][bcap] = int(lim)
            st.markdown("**Disponibilidade semanal — São Luís e Imperatriz**")
            semana_ref = st.date_input("Semana de referência (qualquer dia):", datetime.date.today(), key="semana_ref")
            segunda_ref = semana_ref - datetime.timedelta(days=semana_ref.weekday())
            cols_disp = st.columns(2)
            for j, bdisp in enumerate(["Banca São Luís", "Banca Imperatriz"]):
                with cols_disp[j]:
                    chave_disp = f"{bdisp}|{segunda_ref.isoformat()}"
                    valor_disp = st.number_input(f"{bdisp} — disponível na semana de {segunda_ref.strftime('%d/%m/%Y')}", min_value=0, max_value=100, value=int(st.session_state["disponibilidade_semanal"].get(chave_disp, st.session_state["capacidade_bancas"].get(bdisp, 0))), key=f"disp_{bdisp}_{segunda_ref.isoformat()}")
                    st.session_state["disponibilidade_semanal"][chave_disp] = int(valor_disp)

        if st.button("💾 Registrar Viagem"):
            erros = []
            if d_fim < d_inicio:
                erros.append("A data final não pode ser anterior à data de início.")
            datas_feriado_viagem = []
            chave_fer = f"feriados_{banca_v}_{destino_v}"
            texto_feriados = st.session_state.get("feriados_locais_dict", {}).get(chave_fer, "")
            for linha_feriado in str(texto_feriados).splitlines():
                dt_fer = extrair_data_feriado(linha_feriado.strip(), d_inicio.year)
                if dt_fer and d_inicio <= dt_fer <= d_fim:
                    datas_feriado_viagem.append(dt_fer)
            if datas_feriado_viagem:
                erros.append("A viagem atravessa feriado(s) municipal(is): " + ", ".join(d.strftime("%d/%m/%Y") for d in sorted(set(datas_feriado_viagem))))

            conflitos = conflitos_de_horario_rota(
                banca_v, destino_v, d_inicio, d_fim, turno_v, numero_banca_itinerante_v, apoio_conjunto
            )
            if conflitos:
                erros.append("Existe outro trecho da mesma banca itinerante no mesmo período dentro do intervalo informado.")

            bancas_validar = [banca_v]
            if apoio_conjunto:
                bancas_validar = ["Banca Caxias", "Banca Timon"]
            data_cursor = d_inicio
            while controle_capacidade_ativo and data_cursor <= d_fim:
                if data_cursor.weekday() < 5:
                    for bcap in bancas_validar:
                        qtd_nova = int(num_exam_caxias if bcap == "Banca Caxias" and apoio_conjunto else num_exam_timon if bcap == "Banca Timon" and apoio_conjunto else num_exam_v)
                        picos = []
                        for periodo_cap in ["Manhã", "Tarde"]:
                            por_equipe = {}
                            for v_exist in st.session_state["viagens_registradas"]:
                                if banca_vinculada(v_exist, bcap) and viagem_cobre_periodo(v_exist, data_cursor, periodo_cap):
                                    equipe = int(v_exist.get("Numero Banca Itinerante", 0))
                                    por_equipe[equipe] = max(por_equipe.get(equipe, 0), examinadores_da_banca_na_viagem(v_exist, bcap))
                            equipe_nova = int(numero_banca_itinerante_v)
                            if turno_v == "Dia inteiro" or turno_v == periodo_cap:
                                por_equipe[equipe_nova] = max(por_equipe.get(equipe_nova, 0), qtd_nova)
                            picos.append(sum(por_equipe.values()))
                        pico_fora = max(picos or [0])
                        limite = int(st.session_state["limite_fora_sede_bancas"].get(bcap, st.session_state["capacidade_bancas"].get(bcap, 0)))
                        disponivel_semana = disponibilidade_semanal_banca(bcap, data_cursor)
                        teto = min(limite, disponivel_semana)
                        if pico_fora > teto:
                            erros.append(f"{bcap}: deslocamento simultâneo de {pico_fora} examinadores excede o limite manual de {teto} na semana de {data_cursor.strftime('%d/%m/%Y')}.")
                            break
                data_cursor += datetime.timedelta(days=1)
                if erros and any("capacidade semanal" in e for e in erros):
                    break

            if erros:
                for erro in erros:
                    st.error(erro)
                if datas_feriado_viagem:
                    st.warning("O feriado foi tratado como alerta operacional.")
            else:
                registro = {
                    "Banca": banca_v,
                    "Numero Banca Itinerante": numero_banca_itinerante_v,
                    "Apoio Conjunto": apoio_conjunto,
                    "Destino": destino_v,
                    "Data Inicio": d_inicio,
                    "Data Fim": d_fim,
                    "Turno": turno_v,
                    "Examinadores": int(num_exam_v),
                    "Observações": obs_v,
                }
                if apoio_conjunto:
                    registro["Examinadores Caxias"] = int(num_exam_caxias)
                    registro["Examinadores Timon"] = int(num_exam_timon)
                    registro["Examinadores"] = int(num_exam_caxias + num_exam_timon)
                st.session_state["viagens_registradas"].append(registro)
                st.success(
                    f"Trecho para **{destino_v}** registrado na Banca {numero_banca_itinerante_v:02d}. "
                    f"Período: {turno_v}."
                )

with v_col2:
    st.markdown("#### 📋 Viagens Programadas")

    if "viagens_registradas" not in st.session_state:
        st.session_state["viagens_registradas"] = []

    if not st.session_state["viagens_registradas"]:
        st.warning("Nenhuma viagem cadastrada até o momento.")
    else:
        import pandas as pd

        # Converte as viagens salvas em DataFrame
        df_viagens = pd.DataFrame(st.session_state["viagens_registradas"])

        st.caption("💡 **Dica:** Clique em qualquer célula abaixo para editar datas, bancas ou destinos diretamente.")

        # Tabela totalmente editável em tempo real
        df_editado = st.data_editor(
            df_viagens,
            num_rows="dynamic", # Permite adicionar ou excluir linhas na tabela
            use_container_width=True,
            hide_index=True,
            key="editor_viagens_dinamico"
        )

        # Sincroniza as edições feitas na tela de volta para o sistema
        st.session_state["viagens_registradas"] = df_editado.to_dict("records")

# --- ABA: GERENCIAR HORÁRIOS ---
with aba_horarios:
    st.markdown("### ⏰ Cadastro e Edição de Horários das Turmas")
    col1, col2 = st.columns(2)
    with col1:
        novo_h = st.text_input(
            "Adicionar novo horário (Formato HH:MM):", placeholder="Ex: 07:30"
        )
        if st.button("Adicionar Horário"):
            if novo_h and novo_h not in st.session_state["lista_horarios"]:
                st.session_state["lista_horarios"].append(novo_h)
                st.session_state["lista_horarios"].sort()
                st.success(f"Horário {novo_h} adicionado com sucesso!")

    with col2:
        h_remover = st.selectbox(
            "Remover horário existente:", st.session_state["lista_horarios"]
        )
        if st.button("Remover Horário"):
            st.session_state["lista_horarios"].remove(h_remover)
            st.warning(f"Horário {h_remover} removido!")

# --- ABA: MONTAR CALENDÁRIO ---
with aba_cal:
    st.sidebar.header("⚙️ Parâmetros do Calendário")

    banca_sel = st.sidebar.selectbox(
        "Banca Principal:", list(st.session_state["bancas_config"].keys())
    )
    locais_disp = st.session_state["bancas_config"][banca_sel]
    local_sel = st.sidebar.selectbox("Local de Atendimento:", locais_disp)

    col_ano, col_mes = st.sidebar.columns(2)
    with col_ano:
        ano_sel = st.sidebar.selectbox("Ano:", ANOS_LISTA, index=0)
    with col_mes:
        mes_nome = st.sidebar.selectbox(
            "Mês:", MESES_LISTA, index=10
        )
        mes_num = MESES_LISTA.index(mes_nome) + 1

    st.sidebar.markdown("---")
    st.sidebar.subheader("🛡️ Efetivo Diário da Banca")
    efetivo_diario_banca = st.sidebar.number_input(
        f"Efetivo Máximo Diário ({banca_sel}):",
        min_value=1,
        max_value=200,
        value=30,
    )

    st.sidebar.subheader("📅 Dias de Atendimento na Semana")
    dias_semana_opcoes = DIAS_SEMANA_OPCOES
    dias_permitidos = st.sidebar.multiselect(
        "Selecione os dias de exame desta localidade:",
        options=dias_semana_opcoes,
        default=["Segunda", "Terça", "Quarta", "Quinta", "Sexta"],
    )

    st.session_state["dias_permitidos_dict"][f"{banca_sel}_{local_sel}"] = (
        dias_permitidos
    )

    st.sidebar.subheader("👨‍⚖️ SUGESTÃO INICIAL DE EXAMINADORES")
    def_ex_m = st.sidebar.number_input(
        "Padrão Inicial Manhã:", min_value=0, max_value=50, value=4
    )
    def_ex_t = st.sidebar.number_input(
        "Padrão Inicial Tarde:", min_value=0, max_value=50, value=4
    )

    st.sidebar.subheader("🏙️ Feriados Municipais / Locais")
    chave_feriado_loc = f"feriados_{banca_sel}_{local_sel}"
    val_padrao = st.session_state["feriados_locais_dict"].get(
        chave_feriado_loc, ""
    )

    feriados_txt = st.sidebar.text_area(
        "Digite um por linha (Ex: DD/MM - Motivo):",
        value=val_padrao,
        help="Digite datas no formato DD/MM ou DD/MM/AAAA",
        key=f"txt_{chave_feriado_loc}",
    )
    st.session_state["feriados_locais_dict"][chave_feriado_loc] = feriados_txt

    try:
        feriados_oficiais = holidays.Brazil(subdiv="MA", years=ano_sel)
    except TypeError:
        feriados_oficiais = holidays.BR(state="MA", years=ano_sel)
    datas_feriados_set = set(feriados_oficiais.keys())

    if feriados_txt.strip():
        for l in feriados_txt.strip().split("\n"):
            dt_parse = extrair_data_feriado(l, ano_sel)
            if dt_parse:
                datas_feriados_set.add(dt_parse)

    num_dias = calendar.monthrange(ano_sel, mes_num)[1]
    registros = []
    cols_vagas = [
        "Cat A",
        "Cat B",
        "Cat C",
        "Cat D",
        "Cat E",
        "PCD A",
        "PCD B",
        "PCD C",
        "PCD D",
        "PCD E",
    ]
    datas_disponiveis_mes = []

    is_banca_regional = banca_sel not in ["Banca São Luís", "Banca Imperatriz"]
    is_sede = local_sel == obter_sede(banca_sel)

    viagens_da_localidade = [
        v
        for v in st.session_state["viagens_registradas"]
        if (
            v["Banca"] == banca_sel
            or (
                v.get("Apoio Conjunto")
                and banca_sel in ["Banca Caxias", "Banca Timon"]
            )
        )
        and v["Destino"] == local_sel
    ]
    tem_viagem_para_local = len(viagens_da_localidade) > 0

    for dia in range(1, num_dias + 1):
        data_at = datetime.date(ano_sel, mes_num, dia)
        w_day = data_at.weekday()
        dia_semana_str = dias_semana_opcoes[w_day]

        if dia_semana_str not in dias_permitidos:
            continue

        data_fmt = data_at.strftime("%d/%m/%Y")
        if data_fmt not in datas_disponiveis_mes:
            datas_disponiveis_mes.append(data_fmt)

        is_feriado = data_at in datas_feriados_set

        viagem_local_data = None
        banca_em_viagem_geral = False

        for h in st.session_state["lista_horarios"]:
            periodo_h = periodo_do_horario(h)
            viagem_h = next(
                (v for v in viagens_da_localidade if viagem_cobre_periodo(v, data_at, periodo_h)),
                None,
            )
            viagem_banca_h = viagens_ativas_no_periodo(banca_sel, data_at, periodo_h)
            if is_feriado:
                status_ini_linha = "Feriado / Sem Atendimento"
                ex_m_linha = ex_t_linha = 0
            elif viagem_h:
                ex_viagem = examinadores_da_banca_na_viagem(viagem_h, banca_sel)
                status_ini_linha = "Disponível"
                ex_m_linha = ex_viagem if periodo_h == "Manhã" else 0
                ex_t_linha = ex_viagem if periodo_h == "Tarde" else 0
            elif tem_viagem_para_local:
                status_ini_linha = "Indisponível"
                ex_m_linha = ex_t_linha = 0
            elif is_banca_regional and is_sede and viagem_banca_h:
                status_ini_linha = "Indisponível"
                ex_m_linha = ex_t_linha = 0
            else:
                status_ini_linha = "Disponível"
                ex_m_linha = def_ex_m if periodo_h == "Manhã" else 0
                ex_t_linha = def_ex_t if periodo_h == "Tarde" else 0
            row = {
                "Data": data_fmt,
                "Dia da Semana": dia_semana_str,
                "Status": status_ini_linha,
                "Exam. M": ex_m_linha if status_ini_linha == "Disponível" else 0,
                "Exam. T": ex_t_linha if status_ini_linha == "Disponível" else 0,
                "Horário": h,
            }
            for c in cols_vagas:
                row[c] = 0
            registros.append(row)

    df_base = pd.DataFrame(registros)

    chave_guardar = f"{banca_sel}_{mes_nome}_{ano_sel}_{local_sel}"

    if chave_guardar in st.session_state["historico_localidades"]:
        df_antigo = st.session_state["historico_localidades"][chave_guardar]
        df_antigo_idx = df_antigo.set_index(["Data", "Horário"])

        registros_merged = []
        for _, row_base in df_base.iterrows():
            chave_linha = (row_base["Data"], row_base["Horário"])
            if chave_linha in df_antigo_idx.index:
                row_antiga = df_antigo_idx.loc[chave_linha]
                status_estrutural_bloqueado = row_base["Status"] in [
                    "Feriado / Sem Atendimento",
                    "Indisponível",
                ]
                if status_estrutural_bloqueado:
                    registros_merged.append(row_base.to_dict())
                else:
                    row_final = row_base.to_dict()
                    row_final["Status"] = row_antiga["Status"]
                    row_final["Exam. M"] = row_antiga["Exam. M"]
                    row_final["Exam. T"] = row_antiga["Exam. T"]
                    for c in cols_vagas:
                        row_final[c] = row_antiga[c]
                    registros_merged.append(row_final)
            else:
                registros_merged.append(row_base.to_dict())

        df_completo = pd.DataFrame(registros_merged)
    else:
        df_completo = df_base.copy()

    # --- COMPONENTE VISUAL DO CALENDÁRIO ---
    st.markdown("### 📅 Visualização em Calendário Interativo")
    if HAS_CALENDAR_COMPONENT:
        events = []
        # Renderiza viagens cadastradas da banca selecionada
        for v in st.session_state["viagens_registradas"]:
            if v["Banca"] == banca_sel or (v.get("Apoio Conjunto") and banca_sel in ["Banca Caxias", "Banca Timon"]):
                events.append({
                    "title": f"🚍 Viagem: {v['Destino']} ({v['Examinadores']} Ex.)",
                    "start": v["Data Inicio"].isoformat(),
                    "end": (v["Data Fim"] + datetime.timedelta(days=1)).isoformat(),
                    "color": "#2B6CB0" if v["Destino"] == local_sel else "#718096",
                    "allDay": True,
                })
        
        # Renderiza feriados
        for f_date in datas_feriados_set:
            if f_date.month == mes_num and f_date.year == ano_sel:
                events.append({
                    "title": "🔴 Feriado / Sem Atendimento",
                    "start": f_date.isoformat(),
                    "color": "#E53E3E",
                    "allDay": True,
                })

        calendar_options = {
            "editable": True,
            "selectable": True,
            "headerToolbar": {
                "left": "prev,next today",
                "center": "title",
                "right": "dayGridMonth,timeGridWeek"
            },
            "initialDate": f"{ano_sel}-{mes_num:02d}-01",
            "locale": "pt-br",
        }
        
        cal_data = st_calendar(events=events, options=calendar_options, key=f"cal_visual_{chave_guardar}")
        
        # Sincronização básica se o usuário mover um evento no calendário
        if cal_data.get("eventChange"):
            event_changed = cal_data["eventChange"]["event"]
            title = event_changed.get("title", "")
            if title.startswith("🚍 Viagem:"):
                dest_match = re.search(r"Viagem:\s*([^(]+)", title)
                if dest_match:
                    dest_nome = dest_match.group(1).strip()
                    try:
                        n_start = datetime.date.fromisoformat(event_changed["start"].split("T")[0])
                        n_end = datetime.date.fromisoformat(event_changed["end"].split("T")[0]) - datetime.timedelta(days=1)
                        if n_end < n_start:
                            n_end = n_start
                        
                        for v in st.session_state["viagens_registradas"]:
                            if v["Banca"] == banca_sel and v["Destino"] == dest_nome:
                                v["Data Inicio"] = n_start
                                v["Data Fim"] = n_end
                                st.toast(f"Datas da viagem para {dest_nome} atualizadas via calendário!")
                                st.rerun()
                                break
                    except Exception:
                        pass
    else:
        st.info("💡 Para ativar a interface gráfica e arrastar viagens diretamente na tela, adicione `streamlit-calendar` ao seu `requirements.txt`.")

    st.markdown("### 📝 Lançamento e Edição de Vagas por Horário")
    if tem_viagem_para_local:
        v_ref = viagens_da_localidade[0]
        st.success(
            f"🚍 **Viagem Detectada**: {local_sel} tem viagem registrada de"
            f" **{v_ref['Data Inicio'].strftime('%d/%m/%Y')}** a"
            f" **{v_ref['Data Fim'].strftime('%d/%m/%Y')}** com"
            f" **{v_ref['Examinadores']} examinadores**. As datas foram ativadas"
            " automaticamente abaixo!"
        )
    elif is_banca_regional and is_sede:
        st.info(
            f"📍 Configurando **{local_sel}** ({banca_sel}). As datas em que a equipe"
            " estiver em viagem itinerante ficam marcadas como Indisponível na Sede."
        )
    else:
        st.info(f"📍 Configurando calendário para **{local_sel}** ({banca_sel}).")

    col_filtro1, col_filtro2 = st.columns([1, 2])

    dt_inicio_mes = datetime.date(ano_sel, mes_num, 1)
    dt_fim_mes = datetime.date(ano_sel, mes_num, num_dias)

    with col_filtro1:
        periodo_selecionado = st.date_input(
            "🔍 Filtrar por Período de Atendimento:",
            value=(dt_inicio_mes, dt_fim_mes),
            min_value=dt_inicio_mes,
            max_value=dt_fim_mes,
            format="DD/MM/YYYY",
            key=f"filtro_periodo_{chave_guardar}"
        )

    with col_filtro2:
        datas_opcoes = sorted(list(df_completo["Data"].unique()), key=lambda x: datetime.datetime.strptime(x, "%d/%m/%Y"))
        datas_filtradas_select = st.multiselect(
            "📌 Ou Selecione Dias Específicos do Mês:",
            options=datas_opcoes,
            default=[],
            placeholder="Selecione um ou mais dias...",
            key=f"filtro_datas_mult_{chave_guardar}"
        )

    if datas_filtradas_select:
        df_exibicao = df_completo[df_completo["Data"].isin(datas_filtradas_select)].copy()
    elif isinstance(periodo_selecionado, (tuple, list)) and len(periodo_selecionado) == 2:
        p_ini, p_fim = periodo_selecionado
        df_completo["_dt_temp"] = pd.to_datetime(df_completo["Data"], format="%d/%m/%Y").dt.date
        df_exibicao = df_completo[(df_completo["_dt_temp"] >= p_ini) & (df_completo["_dt_temp"] <= p_fim)].drop(columns=["_dt_temp"]).copy()
    else:
        df_exibicao = df_completo.copy()

    opcoes_status = ["Disponível", "Indisponível", "Feriado / Sem Atendimento"]

    config_cols = {
        "Data": st.column_config.TextColumn("Data", disabled=True),
        "Dia da Semana": st.column_config.TextColumn("Dia", disabled=True),
        "Status": st.column_config.SelectboxColumn(
            "Status", options=opcoes_status, required=True
        ),
        "Exam. M": st.column_config.NumberColumn(
            "Exam. M", min_value=0, max_value=50, step=1
        ),
        "Exam. T": st.column_config.NumberColumn(
            "Exam. T", min_value=0, max_value=50, step=1
        ),
        "Horário": st.column_config.TextColumn("Horário", disabled=True),
    }

    for c in cols_vagas:
        config_cols[c] = st.column_config.NumberColumn(
            c, min_value=0, step=1, default=0
        )

    chave_tabela_estavel = (
        f"editor_tabela_{banca_sel}_{local_sel}_{mes_nome}_{ano_sel}"
    )
    chave_assinatura_filtro = f"{chave_tabela_estavel}__ultimo_filtro"
    assinatura_filtro_atual = (
        tuple(sorted(datas_filtradas_select))
        if datas_filtradas_select
        else str(periodo_selecionado)
    )
    if st.session_state.get(chave_assinatura_filtro) != assinatura_filtro_atual:
        st.session_state.pop(chave_tabela_estavel, None)
        st.session_state[chave_assinatura_filtro] = assinatura_filtro_atual

    df_editado_filtrado = st.data_editor(
        df_exibicao,
        column_config=config_cols,
        use_container_width=True,
        num_rows="fixed",
        height=450,
        key=chave_tabela_estavel,
    )

    df_completo.update(df_editado_filtrado)

    def recalcular_linha(row):
        if row["Status"] in ["Feriado / Sem Atendimento", "Indisponível"]:
            return 0
        return sum(row[c] for c in cols_vagas)

    df_completo["Total"] = df_completo.apply(recalcular_linha, axis=1)
    st.session_state["historico_localidades"][chave_guardar] = df_completo

    # MONITOR DE EFETIVO
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Consulta de Efetivo por Data")

    if datas_disponiveis_mes:
        data_consulta = st.sidebar.selectbox(
            "Escolha a Data para checar o total:", datas_disponiveis_mes
        )

        detalhes_dia = []
        total_m_dia = 0
        total_t_dia = 0

        prefixo_busca = f"{banca_sel}_{mes_nome}_{ano_sel}_"
        for k_loc, df_loc in st.session_state["historico_localidades"].items():
            if k_loc.startswith(prefixo_busca):
                local_nome = k_loc.replace(prefixo_busca, "")
                df_dia = df_loc[
                    (df_loc["Data"] == data_consulta) & (df_loc["Status"] == "Disponível")
                ]

                if not df_dia.empty:
                    data_obj = datetime.datetime.strptime(data_consulta, "%d/%m/%Y").date()
                    viagem_m = next(
                        (v for v in st.session_state["viagens_registradas"]
                         if banca_vinculada(v, banca_sel)
                         and v.get("Destino") == local_nome
                         and viagem_cobre_periodo(v, data_obj, "Manhã")),
                        None,
                    )
                    viagem_t = next(
                        (v for v in st.session_state["viagens_registradas"]
                         if banca_vinculada(v, banca_sel)
                         and v.get("Destino") == local_nome
                         and viagem_cobre_periodo(v, data_obj, "Tarde")),
                        None,
                    )

                    ex_m_loc = 0 if viagem_m else int(df_dia["Exam. M"].max())
                    ex_t_loc = 0 if viagem_t else int(df_dia["Exam. T"].max())
                    total_m_dia += ex_m_loc
                    total_t_dia += ex_t_loc
                    if ex_m_loc or ex_t_loc:
                        detalhes_dia.append(
                            {"Localidade": local_nome, "Manhã": ex_m_loc, "Tarde": ex_t_loc}
                        )

        equipes_manha = {}
        equipes_tarde = {}
        data_obj = datetime.datetime.strptime(data_consulta, "%d/%m/%Y").date()
        for v in st.session_state["viagens_registradas"]:
            if not banca_vinculada(v, banca_sel):
                continue
            equipe = int(v.get("Numero Banca Itinerante", 0))
            qtd = examinadores_da_banca_na_viagem(v, banca_sel)
            if qtd <= 0:
                continue
            if viagem_cobre_periodo(v, data_obj, "Manhã"):
                equipes_manha[equipe] = max(equipes_manha.get(equipe, 0), qtd)
            if viagem_cobre_periodo(v, data_obj, "Tarde"):
                equipes_tarde[equipe] = max(equipes_tarde.get(equipe, 0), qtd)

        total_m_dia += sum(equipes_manha.values())
        total_t_dia += sum(equipes_tarde.values())

        for equipe, qtd in equipes_manha.items():
            detalhes_dia.append({"Localidade": f"Banca itinerante {equipe} (Viagem)", "Manhã": qtd, "Tarde": 0})
        for equipe, qtd in equipes_tarde.items():
            detalhes_dia.append({"Localidade": f"Banca itinerante {equipe} (Viagem)", "Manhã": 0, "Tarde": qtd})

        max_uso_dia = max(total_m_dia, total_t_dia)
        saldo_dia = efetivo_diario_banca - max_uso_dia

        st.sidebar.markdown(f"**Escala do Dia:** `{data_consulta}`")
        st.sidebar.metric(
            "Efetivo Usado no Dia", f"{max_uso_dia} / {efetivo_diario_banca}"
        )
        st.sidebar.metric("Saldo Restante", f"{saldo_dia}")

        if max_uso_dia > efetivo_diario_banca:
            st.sidebar.error(
                "⚠️ **ESTOURO DE EFETIVO!** Excesso de"
                f" {max_uso_dia - efetivo_diario_banca} examinadores."
            )

        if detalhes_dia:
            st.sidebar.markdown("**Distribuição por Localidade nesta Data:**")
            df_detalhes = pd.DataFrame(detalhes_dia)
            st.sidebar.dataframe(
                df_detalhes, use_container_width=True, hide_index=True
            )
        else:
            st.sidebar.info("Sem exames ou viagens nesta data.")

    # INDICADORES INFERIORES
    totais_por_data = {}
    prefixo_banca_mes = f"{banca_sel}_{mes_nome}_{ano_sel}_"

    for k_loc, df_loc in st.session_state["historico_localidades"].items():
        if k_loc.startswith(prefixo_banca_mes):
            df_disp = df_loc[df_loc["Status"] == "Disponível"]
            if not df_disp.empty:
                for data_str, df_grupo_data in df_disp.groupby("Data"):
                    ex_m_loc = df_grupo_data["Exam. M"].max()
                    ex_t_loc = df_grupo_data["Exam. T"].max()

                    if data_str not in totais_por_data:
                        totais_por_data[data_str] = {"M": 0, "T": 0}
                    totais_por_data[data_str]["M"] += ex_m_loc
                    totais_por_data[data_str]["T"] += ex_t_loc

    for dia_num in range(1, num_dias + 1):
        data_obj = datetime.date(ano_sel, mes_num, dia_num)
        data_str = data_obj.strftime("%d/%m/%Y")
        equipes_manha = {}
        equipes_tarde = {}
        for v in st.session_state["viagens_registradas"]:
            if not banca_vinculada(v, banca_sel):
                continue
            qtd = examinadores_da_banca_na_viagem(v, banca_sel)
            equipe = int(v.get("Numero Banca Itinerante", 0))
            if qtd <= 0:
                continue
            if viagem_cobre_periodo(v, data_obj, "Manhã"):
                equipes_manha[equipe] = max(equipes_manha.get(equipe, 0), qtd)
            if viagem_cobre_periodo(v, data_obj, "Tarde"):
                equipes_tarde[equipe] = max(equipes_tarde.get(equipe, 0), qtd)
        if equipes_manha or equipes_tarde:
            if data_str not in totais_por_data:
                totais_por_data[data_str] = {"M": 0, "T": 0}
            totais_por_data[data_str]["M"] += sum(equipes_manha.values())
            totais_por_data[data_str]["T"] += sum(equipes_tarde.values())

    picos_diarios = (
        [max(v["M"], v["T"]) for v in totais_por_data.values()]
        if totais_por_data
        else [max(def_ex_m, def_ex_t)]
    )
    pico_demanda_diaria = max(picos_diarios)

    st.markdown("---")
    m_col1, m_col2, m_col3 = st.columns(3)
    with m_col1:
        st.metric("Total Vagas Localidade", f"{int(df_completo['Total'].sum())} vagas")
    with m_col2:
        st.metric(
            "Maior Pico Diário no Mês",
            f"{int(pico_demanda_diaria)} / {efetivo_diario_banca} examinadores",
        )
    with m_col3:
        saldo_diario = efetivo_diario_banca - pico_demanda_diaria
        st.metric("Menor Folga Diária", f"{int(saldo_diario)} examinadores")

    if pico_demanda_diaria > efetivo_diario_banca:
        st.error(
            "🚨 **BLOQUEIO DE SEGURANÇA**: Excesso de examinadores escalados em"
            " relação ao limite diário disponível."
        )
    else:
        pdf_bytes = gerar_pdf_oficial_enxuto(
            df_completo, banca_sel, local_sel, mes_nome, ano_sel
        )
        st.download_button(
            label="🖨️ Baixar PDF Enxuto desta Localidade",
            data=pdf_bytes,
            file_name=f"Calendario_{banca_sel}_{local_sel}_{mes_nome}_{ano_sel}.pdf",
            mime="application/pdf",
        )

# --- ABA: DASHBOARD E RELATÓRIO CONSOLIDADO ---
with aba_relatorio:
    st.markdown("### 📊 Dashboard Executivo de Oferta de Vagas")

    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    with f_col1:
        mes_rel_sel = st.selectbox("🗓️ Mês:", MESES_LISTA, index=5, key="rel_mes")
    with f_col2:
        ano_rel_sel = st.selectbox("📅 Ano:", [str(a) for a in ANOS_LISTA], key="rel_ano")
    with f_col3:
        bancas_todas = list(st.session_state["bancas_config"].keys())
        banca_rel_sel = st.selectbox("🏛️ Banca:", ["Todas"] + bancas_todas, key="rel_banca")
    with f_col4:
        if banca_rel_sel != "Todas":
            locais_filtro = ["Todas as Localidades"] + st.session_state["bancas_config"].get(banca_rel_sel, [])
        else:
            todos_locais = set()
            for l_list in st.session_state["bancas_config"].values():
                todos_locais.update(l_list)
            locais_filtro = ["Todas as Localidades"] + sorted(list(todos_locais))
        local_rel_sel = st.selectbox("📍 Localidade:", locais_filtro, key="rel_local")

    st.markdown("---")

    resumo_locais = []
    cols_cat = [
        "Cat A",
        "Cat B",
        "Cat C",
        "Cat D",
        "Cat E",
        "PCD A",
        "PCD B",
        "PCD C",
        "PCD D",
        "PCD E",
    ]

    for chave, df_loc in st.session_state["historico_localidades"].items():
        partes = chave.split("_")
        if len(partes) < 4:
            continue
        b_nome = partes[0]
        m_nome = partes[1]
        a_nome = partes[2]
        l_nome = "_".join(partes[3:])

        if m_nome != mes_rel_sel or a_nome != str(ano_rel_sel):
            continue
        if banca_rel_sel != "Todas" and b_nome != banca_rel_sel:
            continue
        if local_rel_sel != "Todas as Localidades" and l_nome != local_rel_sel:
            continue

        df_vagas = df_loc[
            (df_loc["Status"] == "Disponível") & (df_loc["Total"] > 0)
        ].copy()
        total_vagas_loc = df_vagas["Total"].sum()

        ex_dia = 0
        datas_exames_str = "-"

        if not df_vagas.empty:
            grp_ex = df_vagas.groupby("Data")[["Exam. M", "Exam. T"]].max()
            ex_dia = grp_ex.apply(
                lambda r: max(r["Exam. M"], r["Exam. T"]), axis=1
            ).max()

            dias_unicos = sorted(
                list(set([int(d.split("/")[0]) for d in df_vagas["Data"].unique()]))
            )
            datas_exames_str = formatar_datas_exames(dias_unicos)

        linha_resumo = {
            "Banca": b_nome,
            "Localidade": l_nome,
            "Mês/Ano": f"{m_nome}/{a_nome}",
            "Dias com Exame": df_vagas["Data"].nunique() if not df_vagas.empty else 0,
            "Datas dos exames": datas_exames_str,
            "Pico Exam./Dia": int(ex_dia),
            "Total Vagas": int(total_vagas_loc),
        }
        for c in cols_cat:
            linha_resumo[c] = (
                int(df_vagas[c].sum()) if c in df_vagas.columns else 0
            )

        resumo_locais.append(linha_resumo)

    if resumo_locais:
        df_resumo_geral = pd.DataFrame(resumo_locais)

        st.markdown("#### 📈 Resumo Geral da Seleção")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)

        tot_vagas = int(df_resumo_geral["Total Vagas"].sum())
        tot_dias = int(df_resumo_geral["Dias com Exame"].sum())
        med_vagas_dia = round(tot_vagas / tot_dias, 1) if tot_dias > 0 else 0
        pico_max = int(df_resumo_geral["Pico Exam./Dia"].max())

        kpi1.metric("🎯 Total Vagas Ofertadas", f"{tot_vagas} vagas")
        kpi2.metric("📅 Total Dias com Exame", f"{tot_dias} dias")
        kpi3.metric("📊 Média Vagas / Dia", f"{med_vagas_dia} vagas")
        kpi4.metric("👨‍⚖️ Maior Pico Examinadores", f"{pico_max} exam.")

        st.markdown("---")
        st.markdown("#### 🏢 Panorama Detalhado por Localidade")
        
        st.dataframe(df_resumo_geral, use_container_width=True, hide_index=True)

        excel_bytes = gerar_excel_modelo_oficial(df_resumo_geral, mes_rel_sel, ano_rel_sel)
        
        st.download_button(
            label="📥 Exportar Relatório Consolidado (Excel Oficial)",
            data=excel_bytes,
            file_name=f"Calendario_Publicacao_{mes_rel_sel.upper()}_{ano_rel_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Nenhum dado cadastrado ou encontrado para os filtros selecionados.")

# --- ABA: GESTÃO DE LOCALIDADES ---
with aba_gestao:
    st.markdown("### ➕ Gestão de Bancas e Localidades")
    st.info("Adicione ou remova Bancas e Municípios de Atendimento da estrutura do sistema.")

    col_g1, col_g2 = st.columns(2)

    with col_g1:
        st.markdown("#### 🏛️ Cadastrar Nova Banca")
        nova_banca = st.text_input("Nome da Nova Banca:", placeholder="Ex: Banca Balsas")
        if st.button("➕ Adicionar Banca"):
            if nova_banca and nova_banca not in st.session_state["bancas_config"]:
                st.session_state["bancas_config"][nova_banca] = []
                st.success(f"Banca **{nova_banca}** criada com sucesso!")
                st.rerun()
            elif nova_banca in st.session_state["bancas_config"]:
                st.warning("Esta banca já existe no sistema.")

    with col_g2:
        st.markdown("#### 📍 Cadastrar Novo Município / Localidade")
        banca_destino = st.selectbox(
            "Selecione a Banca que receberá o local:",
            list(st.session_state["bancas_config"].keys()),
            key="gestao_banca_dest"
        )
        novo_local = st.text_input("Nome do Município/Localidade:", placeholder="Ex: Balsas - Pátio")
        if st.button("➕ Adicionar Localidade"):
            if novo_local:
                if novo_local not in st.session_state["bancas_config"][banca_destino]:
                    st.session_state["bancas_config"][banca_destino].append(novo_local)
                    st.success(f"Localidade **{novo_local}** vinculada à **{banca_destino}** com sucesso!")
                    st.rerun()
                else:
                    st.warning("Esta localidade já está cadastrada nesta banca.")

    st.markdown("---")
    st.markdown("#### 🗑️ Remover Localidade Existente")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        banca_rem = st.selectbox(
            "Selecione a Banca:",
            list(st.session_state["bancas_config"].keys()),
            key="gestao_banca_rem"
        )
    with col_r2:
        local_rem = st.selectbox(
            "Selecione a Localidade para Remover:",
            st.session_state["bancas_config"][banca_rem],
            key="gestao_local_rem"
        )
        if st.button("🗑️ Remover Localidade"):
            st.session_state["bancas_config"][banca_rem].remove(local_rem)
            st.warning(f"Localidade **{local_rem}** removida da **{banca_rem}**.")
            st.rerun()

# --- AUTOSAVE FINAL ---
salvar_autosave_em_disco()
