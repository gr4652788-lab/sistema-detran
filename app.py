"""DETRAN/MA — Sistema de Gestão de Exames Práticos.

Arquivo único, pronto para substituir o app.py do repositório.

Execução:
    pip install -r requirements.txt
    streamlit run app.py

Variáveis de ambiente opcionais:
    DETRAN_DB         caminho do banco SQLite (padrão: dados/detran.db)
    DETRAN_SNAPSHOTS  pasta dos snapshots JSON (padrão: dados/snapshots)
    DETRAN_LOG_LEVEL  nível de log (padrão: INFO)
    DETRAN_SENHA      se definida, exige senha para abrir o sistema
                      (também aceita st.secrets["senha"])

Organização do arquivo:
    1. Configuração e constantes
    2. Regras de negócio (funções puras, sem Streamlit)
       2.1 Motor de capacidade e gerador automático de calendário
    3. Persistência (SQLite, backup, snapshots)
    4. Exportadores (PDF e Excel)
    5. Componentes de interface
    6. Páginas
    7. Aplicação

Versão 5.0
    Novos ajustes da Divisão: efetivo editável (padrão e por dia), gerador
    com período e dias por categoria, quadro só com a região metropolitana e
    a linha "Viagens", postos de sobra montados por último (mínimo de 3
    examinadores), horários de chegada e retorno das viagens, Cat A em
    blocos de 12 por pista, ordem oficial da planilha com PCD em asterisco e
    o cadastro oficial de municípios. Detalhes em docs/AUDITORIA.md.

Versão 4.0
    Implementa a especificação do módulo de agendamento (capacidade por
    categoria, fim de turno, pistas de moto, restrição de turno, horários
    quebrados, gerador automático e validação em vermelho) e corrige as
    falhas apontadas na auditoria da versão 3.1. Rode os testes com:
        python -m pytest test_app.py
"""

from __future__ import annotations

import base64
import calendar
import datetime
import hashlib
import hmac
import html
import math
import inspect
import io
import json
import logging
import os
import re
import sqlite3
import threading
import unicodedata
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable, Sequence

import holidays
import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import Alignment, Border, Font, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

try:  # Componente opcional de calendário interativo.
    from streamlit_calendar import calendar as componente_calendario

    TEM_COMPONENTE_CALENDARIO = True
except ImportError:
    TEM_COMPONENTE_CALENDARIO = False


# ===========================================================================
# 1. CONFIGURAÇÃO E CONSTANTES
# ===========================================================================

SEPARADOR_CHAVE = "||"
SCHEMA_VERSION = 8

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

DIAS_SEMANA_OPCOES = [
    "Segunda",
    "Terça",
    "Quarta",
    "Quinta",
    "Sexta",
    "Sábado",
    "Domingo",
]
DIAS_UTEIS_PADRAO = DIAS_SEMANA_OPCOES[:5]
ANOS_A_FRENTE = 5


def anos_disponiveis(anos_a_frente: int = ANOS_A_FRENTE) -> list[int]:
    """Anos selecionáveis: o anterior (para consulta) e os próximos."""
    ano_atual = datetime.date.today().year
    return list(range(ano_atual - 1, ano_atual + anos_a_frente))


def indice_ano_atual() -> int:
    """Posição do ano corrente em ``anos_disponiveis`` (padrão dos seletores)."""
    return anos_disponiveis().index(datetime.date.today().year)


def indice_mes_atual() -> int:
    return datetime.date.today().month - 1


def numero_do_mes(nome_mes: str) -> int:
    return MESES_LISTA.index(nome_mes) + 1


def normalizar_texto(texto: str) -> str:
    sem_acento = unicodedata.normalize("NFKD", str(texto))
    sem_acento = "".join(c for c in sem_acento if not unicodedata.combining(c))
    return " ".join(sem_acento.lower().split())


def _alfabetica(nomes: Iterable[str]) -> list[str]:
    """Ordem alfabética que ignora acentos e maiúsculas (Açailândia antes de
    Balsas, e não depois de Z)."""
    return sorted(nomes, key=lambda n: (normalizar_texto(n), n))


# Postos da sede de São Luís. Os nomes seguem o cadastro oficial da Divisão
# (versão 5.0); bancos antigos com "São Luís Pátio" etc. são renomeados na
# migração ``_alinhar_cadastro_oficial``.
LOCAL_SL_PATIO = "São Luís - Pátio DETRAN"
LOCAL_SL_CASTELINHO = "São Luís - Área do Castelinho"
LOCAL_SL_COHATRAC = "São Luís - Cohatrac"
LOCAL_SL_CIDADE_OPERARIA = "São Luís - Cidade Operária"

# Municípios atendidos por Timon e Caxias em conjunto: ficam cadastrados nas
# duas bancas, e a viagem é lançada na banca titular (a outra entra em apoio).
_MUNICIPIOS_TIMON_CAXIAS = [
    "Barra do Corda",
    "Buriti Bravo",
    "Coelho Neto",
    "Colinas",
    "Matões",
    "Passagem Franca",
    "Pastos Bons",
    "São João do Sóter",
    "São João dos Patos",
]

# Região metropolitana: postos fixos que aparecem um a um no quadro de
# examinadores. O que não está aqui é atendido por equipe itinerante e entra
# no quadro só como total de "Viagens". Nas bancas que não estão neste
# dicionário, o posto fixo é a sede (primeira localidade da banca).
REGIAO_METROPOLITANA: dict[str, list[str]] = {
    "Banca São Luís": [
        LOCAL_SL_PATIO,
        LOCAL_SL_CASTELINHO,
        LOCAL_SL_COHATRAC,
        LOCAL_SL_CIDADE_OPERARIA,
        "Paço do Lumiar",
        "Raposa",
        "São José de Ribamar",
    ],
    "Banca Imperatriz": ["Imperatriz"],
}

# Postos montados por último: a oferta deles é o que sobra do efetivo do dia
# depois das demais localidades e das viagens.
POSTOS_DE_SOBRA: dict[str, list[str]] = {
    "Banca São Luís": [LOCAL_SL_PATIO, LOCAL_SL_CASTELINHO],
    "Banca Imperatriz": ["Imperatriz"],
}
# Postos de sobra que recebem toda a sobra do dia por padrão (os demais usam
# o número fixo da barra lateral, limitado à sobra, e são gerados antes).
POSTOS_TODA_SOBRA_PADRAO = {LOCAL_SL_PATIO, "Imperatriz"}
# Nunca pode sobrar menos que isto para os postos de sobra no dia.
MINIMO_SOBRA_POSTOS = 3

BANCAS_PADRAO: dict[str, list[str]] = {
    "Banca São Luís": REGIAO_METROPOLITANA["Banca São Luís"]
    + _alfabetica(
        [
            "Arari",
            "Axixá",
            "Barreirinhas",
            "Brejo",
            "Cantanhede",
            "Carutapera",
            "Chapadinha",
            "Codó",
            "Coroatá",
            "Icatu",
            "Itapecuru-Mirim",
            "Pinheiro",
            "Rosário",
            "Santa Rita",
            "São Bento",
            "Turilândia",
            "Tutóia",
            "Viana",
            "Vitória do Mearim",
        ]
    ),
    "Banca Imperatriz": ["Imperatriz"]
    + _alfabetica(
        [
            "Açailândia",
            "Amarante do Maranhão",
            "Balsas",
            "Bom Jesus das Selvas",
            "Buriticupu",
            "Campestre do Maranhão",
            "Cidelândia",
            "Estreito",
            "Governador Edison Lobão",
            "Grajaú",
            "Itinga do Maranhão",
            "João Lisboa",
            "Porto Franco",
        ]
    ),
    "Banca Timon": ["Timon"] + _alfabetica(_MUNICIPIOS_TIMON_CAXIAS),
    "Banca Bacabal": ["Bacabal"]
    + _alfabetica(
        [
            "Alto Alegre",
            "Dom Pedro",
            "Lago da Pedra",
            "Pedreiras",
            "Presidente Dutra",
            "São Mateus do Maranhão",
            "Trizidela do Vale",
            "Vitorino Freire",
        ]
    ),
    "Banca Caxias": ["Caxias"] + _alfabetica(_MUNICIPIOS_TIMON_CAXIAS),
    "Banca Santa Inês": ["Santa Inês"]
    + _alfabetica(["Arari", "Santa Luzia", "Vitória do Mearim"]),
}

# Cadastro semeado pela versão 4 e nomes que mudaram no cadastro oficial.
# Usado só na migração de bancos antigos (``_alinhar_cadastro_oficial``).
RENOMEACOES_V5: dict[str, dict[str, str]] = {
    "Banca São Luís": {
        "São Luís Pátio": LOCAL_SL_PATIO,
        "São Luís Castelinho": LOCAL_SL_CASTELINHO,
        "São Luís Cohatrac": LOCAL_SL_COHATRAC,
        "São Luís Cidade Operária": LOCAL_SL_CIDADE_OPERARIA,
        "Itapecuru-mirim": "Itapecuru-Mirim",
    },
    "Banca Timon": {"Timon - Pátio": "Timon"},
    "Banca Caxias": {"Caxias - Pátio": "Caxias"},
    "Banca Bacabal": {"Bacabal - Pátio": "Bacabal"},
    "Banca Santa Inês": {"Santa Inês - Pátio": "Santa Inês"},
}
# Localidades que a versão 4 semeava e que não estão no cadastro oficial.
# A migração só as remove se nunca receberam vagas nem viagens.
OBSOLETAS_V5: dict[str, list[str]] = {
    "Banca Timon": ["Parnarama"],
    "Banca Caxias": ["Codó - Regional", "Aldeias Altas"],
    "Banca Bacabal": ["Olho d'Água das Cunhãs"],
    "Banca Santa Inês": ["Viana", "Zé Doca", "Monção", "Pindaré-Mirim"],
}

# Ordem da planilha oficial de publicação (Excel).
ORDEM_BANCAS_PUBLICACAO = [
    "Banca São Luís",
    "Banca Imperatriz",
    "Banca Timon",
    "Banca Bacabal",
    "Banca Caxias",
    "Banca Santa Inês",
]
ORDEM_PRIORITARIA_PUBLICACAO: dict[str, list[str]] = {
    "Banca São Luís": REGIAO_METROPOLITANA["Banca São Luís"],
}

# Ordem de apresentação nos relatórios em PDF: primeiro as unidades da sede,
# depois a região metropolitana e, por último, os demais municípios por ordem
# da data do primeiro exame lançado.
ORDEM_PRIORITARIA_PDF: dict[str, list[str]] = {
    "Banca São Luís": [
        LOCAL_SL_PATIO,
        LOCAL_SL_CASTELINHO,
        LOCAL_SL_COHATRAC,
        LOCAL_SL_CIDADE_OPERARIA,
        "Paço do Lumiar",
        "São José de Ribamar",
        "Raposa",
    ],
}

GRUPOS_PDF: dict[str, str] = {
    LOCAL_SL_PATIO: "SEDE — SÃO LUÍS",
    LOCAL_SL_CASTELINHO: "SEDE — SÃO LUÍS",
    LOCAL_SL_COHATRAC: "SEDE — SÃO LUÍS",
    LOCAL_SL_CIDADE_OPERARIA: "SEDE — SÃO LUÍS",
    "Paço do Lumiar": "REGIÃO METROPOLITANA",
    "São José de Ribamar": "REGIÃO METROPOLITANA",
    "Raposa": "REGIÃO METROPOLITANA",
}
GRUPO_INTERIOR = "DEMAIS MUNICÍPIOS"
GRUPO_SEDE_GENERICO = "SEDE DA BANCA"

HORARIOS_PADRAO = [
    "08:30",
    "09:30",
    "10:30",
    "11:30",
    "13:30",
    "14:30",
    "15:30",
    "16:30",
]

CAPACIDADE_BANCAS_PADRAO: dict[str, int] = {
    "Banca São Luís": 34,
    "Banca Imperatriz": 14,
    "Banca Bacabal": 5,
    "Banca Caxias": 4,
    "Banca Timon": 4,
    "Banca Santa Inês": 3,
}
LIMITE_FORA_SEDE_PADRAO: dict[str, int] = {
    "Banca São Luís": 16,
    "Banca Imperatriz": 10,
    "Banca Bacabal": 5,
    "Banca Caxias": 4,
    "Banca Timon": 4,
    "Banca Santa Inês": 3,
}

BANCAS_APOIO_CONJUNTO = ["Banca Caxias", "Banca Timon"]
BANCAS_COM_QUADRO_MATRIZ = ["Banca São Luís", "Banca Imperatriz"]

COLS_CATEGORIA = ["Cat A", "Cat B", "Cat C", "Cat D", "Cat E"]
COLS_PCD = ["PCD A", "PCD B", "PCD C", "PCD D", "PCD E"]
COLS_VAGAS = COLS_CATEGORIA + COLS_PCD
# Coluna interna (oculta no editor): 1 quando o status bloqueado foi posto
# pelo sistema (feriado, viagem, dia fora da grade) e deve se desfazer sozinho.
COL_AUTO = "Bloqueio Auto"

STATUS_DISPONIVEL = "Disponível"
STATUS_INDISPONIVEL = "Indisponível"
STATUS_FERIADO = "Feriado / Sem Atendimento"
OPCOES_STATUS = [STATUS_DISPONIVEL, STATUS_INDISPONIVEL, STATUS_FERIADO]
STATUS_BLOQUEADOS = {STATUS_INDISPONIVEL, STATUS_FERIADO}

TURNO_INTEGRAL = "Dia inteiro"
TURNO_MANHA = "Manhã"
TURNO_TARDE = "Tarde"
TURNO_SEM_ATENDIMENTO = "Não atende"
TURNOS = [TURNO_INTEGRAL, TURNO_MANHA, TURNO_TARDE]
TURNOS_POR_DIA = [TURNO_INTEGRAL, TURNO_MANHA, TURNO_TARDE, TURNO_SEM_ATENDIMENTO]

UF_FERIADOS = "MA"

# Pontos facultativos que a biblioteca ``holidays`` conhece (categoria
# "optional"). Só bloqueiam o calendário os nomes marcados na configuração.
FACULTATIVOS_PADRAO_BLOQUEADOS = ["Carnaval", "Corpus Christi"]

# Bancas regionais: a sede fica indisponível quando a própria banca está em
# viagem (o efetivo é pequeno demais para manter a sede e a itinerante).
BANCAS_REGIONAIS = ["Banca Timon", "Banca Caxias", "Banca Bacabal", "Banca Santa Inês"]

# --- Regras de capacidade (especificação do módulo de agendamento) --------
# A carga é contada em "sextos de examinador" para manter a conta em números
# inteiros: 1 examinador = 6 unidades por horário.
UNIDADES_POR_EXAMINADOR = 6
CATEGORIAS = ["A", "B", "C", "D", "E"]
# Exames por examinador em um horário normal.
EXAMES_POR_EXAMINADOR = {"A": 6, "B": 2, "C": 2, "D": 2, "E": 2}
# Categorias que sofrem o corte de 50% nos horários de fim de turno.
# Cat A ficou de fora por decisão do gestor (vale o exemplo da especificação).
CATEGORIAS_COM_CORTE_FIM_TURNO = {"B", "C", "D", "E"}
CATEGORIAS_SO_MANHA = {"A", "C", "D", "E"}
EXAMINADORES_POR_DUPLA_MOTO = 2
# Horários quebrados: deslocamento em minutos quando a categoria divide o
# horário com a Cat B.
OFFSET_MINUTOS = {"A": 0, "B": 0, "C": 1, "D": 2, "E": 10}
HORARIOS_FIM_TURNO_PADRAO = ["11:30", "16:30"]
PISTAS_MOTO_PADRAO: dict[str, int] = {f"Banca São Luís||{LOCAL_SL_PATIO}": 3}
# Postos em que a Cat A também cai pela metade no fim de turno (no Pátio do
# DETRAN, 11:30 oferta metade dos horários anteriores: 24, 24, 24, 12).
CORTE_CAT_A_FIM_TURNO_PADRAO = {f"Banca São Luís||{LOCAL_SL_PATIO}"}
# Vagas de Cat A saem sempre em blocos fechados: 12 por pista (dupla) no
# horário; 6 no fim de turno dos postos com corte.
VAGAS_CAT_A_POR_DUPLA = EXAMES_POR_EXAMINADOR["A"] * EXAMINADORES_POR_DUPLA_MOTO
# Única coluna PCD que o gerador automático distribui (as demais são manuais).
COL_PCD_GERADA = "PCD B"
DIAS_PCD_PADRAO = ["Quinta"]
JANELA_PADRAO = ("00:00", "23:59")

# Logística padrão atrelada a cada equipe itinerante. Aparece apenas no PDF da
# escala de viagens, conforme solicitado — não há campo para isso na interface.
PREPOSTOS_POR_EQUIPE = 1
VEICULOS_POR_EQUIPE = 1

COR_PRIMARIA = "#1A365D"
COR_SECUNDARIA = "#2B6CB0"
COR_CLARA = "#EDF2F7"
COR_NEUTRA = "#718096"
COR_ALERTA = "#E53E3E"
COR_VIAGEM = "#2F855A"

# Destaque das linhas já lançadas no editor de calendário: verde para linha
# com vagas, cinza para bloqueada (feriado/indisponível).
COR_LANCADO_BG = "#C6F6D5"
COR_LANCADO_TXT = "#22543D"
COR_INATIVO_BG = "#E2E8F0"
COR_INATIVO_TXT = "#4A5568"
# Linha com capacidade estourada ou regra violada.
COR_ERRO_BG = "#FED7D7"
COR_ERRO_TXT = "#9B2C2C"

SITUACAO_OK = "OK"
SITUACAO_ERRO = "ESTOURO"

CSS_CUSTOMIZADO = """
<style>
:root {
    --azul-escuro: #0F2A4A;
    --azul: #1D4E89;
    --azul-claro: #E8EFF8;
    --borda: #D5DCE6;
    --texto-suave: #5B6B80;
}
.block-container { padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1500px; }
#MainMenu, footer { visibility: hidden; }

/* Cabeçalho institucional */
.cabecalho-app {
    display: flex; align-items: center; justify-content: space-between; gap: 16px;
    background: linear-gradient(115deg, var(--azul-escuro) 0%, var(--azul) 62%, #2A64A8 100%);
    border-radius: 14px; padding: 18px 24px; margin-bottom: 14px;
    box-shadow: 0 8px 24px rgba(15, 42, 74, 0.18);
}
.cabecalho-app .marca { display: flex; align-items: center; gap: 16px; }
.cabecalho-app .selo {
    width: 52px; height: 52px; border-radius: 12px; flex: none;
    background: rgba(255, 255, 255, 0.12); border: 1px solid rgba(255, 255, 255, 0.25);
    display: grid; place-items: center; color: #fff; font-weight: 800;
    font-size: 0.78rem; letter-spacing: 0.06em; line-height: 1.05; text-align: center;
}
.cabecalho-app .titulo { color: #fff; font-size: 1.32rem; font-weight: 700; margin: 0; line-height: 1.25; }
.cabecalho-app .subtitulo { color: #C9D7EA; font-size: 0.86rem; margin-top: 2px; }
.cabecalho-app .lado {
    color: #C9D7EA; font-size: 0.8rem; text-align: right; white-space: nowrap;
}
.cabecalho-app .lado b { color: #fff; font-size: 0.95rem; display: block; }

/* Título de página e de seção */
.titulo-pagina { margin: 6px 0 14px 0; }
.titulo-pagina .nome { font-size: 1.45rem; font-weight: 700; color: var(--azul-escuro); line-height: 1.2; }
.titulo-pagina .descricao { color: var(--texto-suave); font-size: 0.92rem; margin-top: 4px; }
.titulo-secao {
    font-size: 1.02rem; font-weight: 700; color: var(--azul-escuro);
    border-left: 4px solid var(--azul); padding: 2px 0 2px 10px; margin: 22px 0 10px 0;
}
.faixa-banca {
    background: var(--azul-claro); color: var(--azul-escuro); border: 1px solid var(--borda);
    padding: 8px 12px; font-weight: 700; border-radius: 8px; font-size: 0.92rem;
    margin: 16px 0 10px 0;
}

/* Indicadores */
div[data-testid="stMetric"] {
    background: #fff; border: 1px solid var(--borda); border-radius: 12px;
    padding: 14px 16px; box-shadow: 0 1px 2px rgba(15, 42, 74, 0.05);
}
div[data-testid="stMetricLabel"] p { color: var(--texto-suave); font-weight: 600; }

/* Navegação principal */
div[data-testid="stSegmentedControl"] { margin-bottom: 6px; }
div[data-testid="stSegmentedControl"] button { font-weight: 600; }

/* Barra lateral */
section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {
    color: var(--azul-escuro); font-size: 0.98rem;
}
.stButton > button, .stDownloadButton > button { font-weight: 600; }
div[data-testid="stExpander"] details { border-radius: 10px; }
</style>
"""


def cabecalho_html() -> str:
    hoje = datetime.date.today()
    return f"""
<div class="cabecalho-app">
  <div class="marca">
    <div class="selo">DETRAN<br>MA</div>
    <div>
      <div class="titulo">Gestão de Exames Práticos</div>
      <div class="subtitulo">Divisão de Exames de Tráfego · calendário de vagas,
      efetivo de examinadores e bancas itinerantes</div>
    </div>
  </div>
  <div class="lado"><b>{DIAS_SEMANA_OPCOES[hoje.weekday()]}, {hoje.strftime('%d/%m/%Y')}</b>
  {MESES_LISTA[hoje.month - 1]} de {hoje.year}</div>
</div>
"""

LOG = logging.getLogger("detran")
Viagem = dict[str, Any]


def _kwargs_largura_total() -> dict[str, Any]:
    """Compatibilidade entre versões do Streamlit.

    ``use_container_width`` foi substituído por ``width="stretch"``. Detectamos
    a assinatura em tempo de execução para o arquivo funcionar tanto em
    instalações antigas quanto nas novas, sem emitir avisos de depreciação.
    """
    try:
        parametros = inspect.signature(st.dataframe).parameters
    except (TypeError, ValueError):
        return {"use_container_width": True}
    if "width" in parametros:
        return {"width": "stretch"}
    return {"use_container_width": True}


LARGURA_TOTAL = _kwargs_largura_total()


# ===========================================================================
# 2. REGRAS DE NEGÓCIO
# ===========================================================================
# Nenhuma função desta seção acessa session_state ou banco: todas recebem os
# dados por argumento, o que as torna testáveis isoladamente.

_PADRAO_ISO = re.compile(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})")
_PADRAO_BR = re.compile(r"^\s*(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?")


def para_data(valor: Any) -> datetime.date | None:
    """Normaliza qualquer representação de data para ``datetime.date``."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime.datetime):
        return valor.date()
    if isinstance(valor, datetime.date):
        return valor
    texto = str(valor).strip()
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.datetime.strptime(texto[:10], formato).date()
        except ValueError:
            continue
    LOG.warning("Não foi possível interpretar a data %r", valor)
    return None


def formatar_br(data: datetime.date | None) -> str:
    return data.strftime("%d/%m/%Y") if data else "-"


def formatar_curto(data: datetime.date | None) -> str:
    return data.strftime("%d/%m") if data else "-"


def semana_chave(data: datetime.date) -> str:
    segunda = data - datetime.timedelta(days=data.weekday())
    return segunda.isoformat()


def nome_dia_semana(data: datetime.date) -> str:
    return DIAS_SEMANA_OPCOES[data.weekday()]


def dias_do_intervalo(
    inicio: datetime.date, fim: datetime.date, apenas_uteis: bool = False
) -> list[datetime.date]:
    dias = []
    cursor = inicio
    while cursor <= fim:
        if not apenas_uteis or cursor.weekday() < 5:
            dias.append(cursor)
        cursor += datetime.timedelta(days=1)
    return dias


def _data_ou_none(ano: int, mes: int, dia: int) -> datetime.date | None:
    try:
        return datetime.date(ano, mes, dia)
    except ValueError:
        return None


@lru_cache(maxsize=64)
def _feriados_publicos(ano: int, uf: str) -> frozenset[datetime.date]:
    try:
        calendario = holidays.Brazil(subdiv=uf, years=ano)
    except TypeError:
        try:
            calendario = holidays.BR(state=uf, years=ano)
        except Exception:
            LOG.exception("Falha ao carregar feriados de %s/%s", uf, ano)
            return frozenset()
    except Exception:
        LOG.exception("Falha ao carregar feriados de %s/%s", uf, ano)
        return frozenset()
    return frozenset(calendario.keys())


@lru_cache(maxsize=64)
def pontos_facultativos(ano: int, uf: str = UF_FERIADOS) -> tuple[tuple[datetime.date, str], ...]:
    """Pontos facultativos conhecidos pela biblioteca (Carnaval, Corpus
    Christi, Quarta de Cinzas etc.). Versões antigas sem categorias devolvem
    vazio, sem quebrar o sistema."""
    try:
        completo = holidays.Brazil(
            subdiv=uf, years=ano, categories=("public", "optional")
        )
        publico = _feriados_publicos(ano, uf)
    except Exception:
        return ()
    return tuple(
        sorted((d, str(n)) for d, n in completo.items() if d not in publico)
    )


def feriados_estaduais(
    ano: int,
    uf: str = UF_FERIADOS,
    facultativos: Iterable[str] | None = None,
) -> frozenset[datetime.date]:
    """Feriados nacionais e estaduais, mais os pontos facultativos cujo nome
    contenha algum dos textos de ``facultativos`` (ex.: "Carnaval")."""
    datas = set(_feriados_publicos(ano, uf))
    nomes = [normalizar_texto(n) for n in (facultativos or []) if str(n).strip()]
    if nomes:
        for data, nome in pontos_facultativos(ano, uf):
            if any(n in normalizar_texto(nome) for n in nomes):
                datas.add(data)
    return frozenset(datas)


def extrair_data_feriado(linha: str, ano_referencia: int) -> datetime.date | None:
    """Lê ``25/12 - Natal``, ``25/12/2027`` ou ``2026-06-15`` e devolve a data."""
    if not linha or not linha.strip():
        return None
    texto = linha.strip()

    iso = _PADRAO_ISO.match(texto)
    if iso:
        return _data_ou_none(*(int(g) for g in iso.groups()))

    br = _PADRAO_BR.match(texto)
    if not br:
        return None
    dia, mes, ano = br.groups()
    if ano is None:
        ano_int = int(ano_referencia)
    else:
        ano_int = int(ano)
        if ano_int < 100:
            ano_int += 2000
    return _data_ou_none(ano_int, int(mes), int(dia))


def datas_de_feriado_do_texto(texto: str, ano_referencia: int) -> set[datetime.date]:
    datas: set[datetime.date] = set()
    for linha in str(texto or "").splitlines():
        data = extrair_data_feriado(linha.strip(), ano_referencia)
        if data:
            datas.add(data)
    return datas


# --- Chaves compostas -------------------------------------------------------
def _sanitizar(texto: str) -> str:
    return str(texto).replace(SEPARADOR_CHAVE, "/")


def chave_historico(banca: str, mes: str, ano: int | str, local: str) -> str:
    return SEPARADOR_CHAVE.join(
        [_sanitizar(banca), _sanitizar(mes), str(ano), _sanitizar(local)]
    )


def desmontar_chave_historico(chave: str) -> tuple[str, str, str, str] | None:
    partes = chave.split(SEPARADOR_CHAVE)
    if len(partes) != 4:
        return None
    return partes[0], partes[1], partes[2], partes[3]


def chave_local(banca: str, local: str) -> str:
    return SEPARADOR_CHAVE.join([_sanitizar(banca), _sanitizar(local)])


def chave_disponibilidade(banca: str, data: datetime.date) -> str:
    return SEPARADOR_CHAVE.join([_sanitizar(banca), semana_chave(data)])


# --- Estrutura organizacional ----------------------------------------------
def obter_sede(bancas_config: dict[str, list[str]], banca: str) -> str | None:
    localidades = bancas_config.get(banca) or []
    return localidades[0] if localidades else None


def todas_localidades(bancas_config: dict[str, list[str]]) -> list[str]:
    encontradas: set[str] = set()
    for lista in bancas_config.values():
        encontradas.update(lista)
    return _alfabetica(encontradas)


def _presentes(
    nomes_referencia: Sequence[str], localidades: Sequence[str]
) -> list[str]:
    """Nomes de ``localidades`` que batem com a referência (ignorando acento e
    caixa), na ordem da referência."""
    por_nome = {normalizar_texto(l): l for l in localidades}
    return [
        por_nome[normalizar_texto(n)]
        for n in nomes_referencia
        if normalizar_texto(n) in por_nome
    ]


def postos_fixos(bancas_config: dict[str, list[str]], banca: str) -> list[str]:
    """Postos fixos da banca (região metropolitana, ou só a sede).

    Os demais municípios da banca são atendidos por equipe itinerante: só têm
    grade nos dias da viagem e aparecem no quadro somados em "Viagens".
    """
    localidades = bancas_config.get(banca) or []
    if banca in REGIAO_METROPOLITANA:
        return _presentes(REGIAO_METROPOLITANA[banca], localidades)
    sede = obter_sede(bancas_config, banca)
    return [sede] if sede else []


def eh_posto_fixo(bancas_config: dict[str, list[str]], banca: str, local: str) -> bool:
    return local in postos_fixos(bancas_config, banca)


def localidades_itinerantes(bancas_config: dict[str, list[str]], banca: str) -> list[str]:
    fixos = set(postos_fixos(bancas_config, banca))
    return [l for l in bancas_config.get(banca) or [] if l not in fixos]


def postos_de_sobra(bancas_config: dict[str, list[str]], banca: str) -> list[str]:
    return _presentes(POSTOS_DE_SOBRA.get(banca, []), bancas_config.get(banca) or [])


def ordenar_localidades_da_banca(banca: str, localidades: Sequence[str]) -> list[str]:
    """Ordem oficial do cadastro: postos da sede/região metropolitana (ou a
    sede) primeiro e o resto em ordem alfabética."""
    prioridade = REGIAO_METROPOLITANA.get(banca)
    if prioridade is None:
        padrao = BANCAS_PADRAO.get(banca)
        prioridade = padrao[:1] if padrao else list(localidades[:1])
    primeiros = _presentes(prioridade, localidades)
    return primeiros + _alfabetica(l for l in localidades if l not in primeiros)


# --- Viagens ----------------------------------------------------------------
def bancas_da_viagem(viagem: Viagem) -> list[str]:
    bancas = []
    titular = viagem.get("Banca")
    if titular:
        bancas.append(titular)
    for banca in viagem.get("Bancas Apoio") or []:
        if banca not in bancas:
            bancas.append(banca)
    return bancas


def banca_vinculada(viagem: Viagem, banca: str) -> bool:
    return banca in bancas_da_viagem(viagem)


def examinadores_da_banca_na_viagem(viagem: Viagem, banca: str) -> int:
    if not banca_vinculada(viagem, banca):
        return 0
    por_banca = viagem.get("Examinadores Por Banca") or {}
    if banca in por_banca:
        try:
            return int(por_banca[banca])
        except (TypeError, ValueError):
            return 0
    if viagem.get("Banca") == banca:
        try:
            return int(viagem.get("Examinadores", 0))
        except (TypeError, ValueError):
            return 0
    return 0


def total_examinadores(viagem: Viagem) -> int:
    por_banca = viagem.get("Examinadores Por Banca") or {}
    if por_banca:
        return sum(int(v or 0) for v in por_banca.values())
    try:
        return int(viagem.get("Examinadores", 0))
    except (TypeError, ValueError):
        return 0


def turno_da_viagem_no_dia(viagem: Viagem, data: datetime.date) -> str:
    """Turno efetivo da viagem em um dia específico.

    O turno geral da viagem vale para todo o período, mas cada dia pode ter
    um ajuste próprio. É o que permite uma mesma equipe atender Pinheiro pela
    manhã e São Bento à tarde no dia 12, sem que o sistema conte duas equipes.
    """
    ajustes = viagem.get("Turnos Por Data") or {}
    especifico = ajustes.get(data.isoformat())
    if especifico:
        return especifico
    return viagem.get("Turno", TURNO_INTEGRAL)


def viagem_cobre_periodo(viagem: Viagem, data: datetime.date, periodo: str) -> bool:
    inicio = para_data(viagem.get("Data Inicio"))
    fim = para_data(viagem.get("Data Fim"))
    if not inicio or not fim or not (inicio <= data <= fim):
        return False
    turno = turno_da_viagem_no_dia(viagem, data)
    if turno == TURNO_SEM_ATENDIMENTO:
        return False
    return turno == TURNO_INTEGRAL or turno == periodo


def viagem_cobre_dia(viagem: Viagem, data: datetime.date) -> bool:
    return viagem_cobre_periodo(viagem, data, TURNO_MANHA) or viagem_cobre_periodo(
        viagem, data, TURNO_TARDE
    )


def dias_efetivos_da_viagem(viagem: Viagem) -> list[datetime.date]:
    """Dias em que a equipe realmente atende, já descontando 'Não atende'."""
    inicio = para_data(viagem.get("Data Inicio"))
    fim = para_data(viagem.get("Data Fim"))
    if not inicio or not fim:
        return []
    return [d for d in dias_do_intervalo(inicio, fim) if viagem_cobre_dia(viagem, d)]


def viagem_presente_no_dia(viagem: Viagem, data: datetime.date) -> bool:
    """A equipe está fora da sede neste dia (atendendo ou em deslocamento).

    É a regra do efetivo: o dia de deslocamento ocupa a equipe inteira, mesmo
    sem exame, e por isso conta nos dois turnos do quadro de examinadores.
    """
    inicio = para_data(viagem.get("Data Inicio"))
    fim = para_data(viagem.get("Data Fim"))
    return bool(inicio and fim and inicio <= data <= fim)


def horarios_limite_da_viagem(viagem: Viagem) -> tuple[str, str]:
    """(primeiro horário de exame no dia de chegada, último no dia de retorno).

    Vazio = sem limite. Servem para não ofertar exame enquanto a equipe ainda
    está na estrada (ex.: Codó, 5 h de viagem: só há prova à tarde no 1º dia).
    """
    return (
        str(viagem.get("Horario Inicio") or "").strip(),
        str(viagem.get("Horario Fim") or "").strip(),
    )


def viagem_cobre_horario(viagem: Viagem, data: datetime.date, horario: str) -> bool:
    """O horário tem exame da equipe: turno do dia e, no primeiro e no último
    dia de atendimento, os horários de chegada e de retorno."""
    if not viagem_cobre_periodo(viagem, data, periodo_do_horario(horario)):
        return False
    minutos = _hora_em_minutos(horario)
    if minutos is None:
        return True
    dias = dias_efetivos_da_viagem(viagem)
    if not dias:
        return False
    primeiro_horario, ultimo_horario = horarios_limite_da_viagem(viagem)
    if data == dias[0] and primeiro_horario:
        limite = _hora_em_minutos(primeiro_horario)
        if limite is not None and minutos < limite:
            return False
    if data == dias[-1] and ultimo_horario:
        limite = _hora_em_minutos(ultimo_horario)
        if limite is not None and minutos > limite:
            return False
    return True


def descricao_turnos(viagem: Viagem) -> str:
    """Texto curto com os dias que fogem do turno geral da viagem."""
    ajustes = viagem.get("Turnos Por Data") or {}
    if not ajustes:
        return ""
    partes = []
    for iso in sorted(ajustes):
        data = para_data(iso)
        turno = ajustes[iso]
        if not data or turno == viagem.get("Turno", TURNO_INTEGRAL):
            continue
        partes.append(f"{formatar_curto(data)}: {turno.lower()}")
    return " · ".join(partes)


def periodo_do_horario(horario: str) -> str:
    try:
        hora, minuto = (int(p) for p in str(horario).split(":")[:2])
    except (TypeError, ValueError):
        LOG.warning("Horário inválido: %r — assumindo Manhã", horario)
        return TURNO_MANHA
    return TURNO_MANHA if (hora, minuto) < (12, 0) else TURNO_TARDE


def viagens_fora_da_sede(
    viagens: Sequence[Viagem], banca: str, data: datetime.date
) -> list[Viagem]:
    """Viagens com examinadores da banca fora da sede no dia (inclui o
    deslocamento)."""
    return [
        v
        for v in viagens
        if banca_vinculada(v, banca)
        and examinadores_da_banca_na_viagem(v, banca) > 0
        and viagem_presente_no_dia(v, data)
    ]


def viagem_do_local_no_dia(
    viagens: Sequence[Viagem], banca: str, local: str, data: datetime.date
) -> Viagem | None:
    for v in viagens:
        if (
            banca_vinculada(v, banca)
            and v.get("Destino") == local
            and viagem_cobre_dia(v, data)
        ):
            return v
    return None


def conflitos_de_rota(
    viagens: Sequence[Viagem],
    banca: str,
    numero_equipe: int,
    inicio: datetime.date,
    fim: datetime.date,
    turno: str,
    turnos_por_data: dict[str, str] | None = None,
    ignorar_id: Any = None,
) -> list[Viagem]:
    """Trechos da mesma equipe que disputam o mesmo turno no mesmo dia.

    A checagem passou a ser dia a dia: antes bastava haver sobreposição de
    intervalo para bloquear, o que impedia justamente o caso de Pinheiro pela
    manhã e São Bento à tarde no mesmo dia.

    A equipe é identificada por (banca titular, número): a equipe 01 de
    Caxias e a equipe 01 de Timon são equipes diferentes, mesmo quando Timon
    entra em apoio na viagem de Caxias.
    """
    candidata = {
        "Data Inicio": inicio,
        "Data Fim": fim,
        "Turno": turno,
        "Turnos Por Data": turnos_por_data or {},
    }
    conflitos: list[Viagem] = []

    for v in viagens:
        if ignorar_id is not None and v.get("id") == ignorar_id:
            continue
        if chave_equipe(v) != (banca, int(numero_equipe)):
            continue
        v_ini = para_data(v.get("Data Inicio"))
        v_fim = para_data(v.get("Data Fim"))
        if not v_ini or not v_fim:
            continue

        for dia in dias_do_intervalo(max(inicio, v_ini), min(fim, v_fim)):
            for periodo in (TURNO_MANHA, TURNO_TARDE):
                if viagem_cobre_periodo(candidata, dia, periodo) and (
                    viagem_cobre_periodo(v, dia, periodo)
                ):
                    conflitos.append(v)
                    break
            else:
                continue
            break
    return conflitos


def chave_equipe(viagem: Viagem) -> tuple[str, int]:
    """Identidade da equipe itinerante: (banca titular, número)."""
    return (
        str(viagem.get("Banca", "")),
        int(viagem.get("Numero Banca Itinerante", 0) or 0),
    )


def numeros_de_equipe(viagens: Sequence[Viagem], banca: str) -> list[int]:
    return sorted(
        {
            int(v.get("Numero Banca Itinerante", 0) or 0)
            for v in viagens
            if v.get("Banca") == banca
        }
        - {0}
    )


def proximo_numero_equipe(viagens: Sequence[Viagem], banca: str) -> int:
    numeros = numeros_de_equipe(viagens, banca)
    return max(numeros) + 1 if numeros else 1


def garantir_numeracao_equipes(viagens: list[Viagem]) -> list[Viagem]:
    contador: dict[str, int] = {}
    for v in viagens:
        banca = v.get("Banca", "")
        numero = v.get("Numero Banca Itinerante")
        if not numero:
            contador[banca] = contador.get(banca, 0) + 1
            v["Numero Banca Itinerante"] = contador[banca]
        else:
            contador[banca] = max(contador.get(banca, 0), int(numero))
    return viagens


# --- Efetivo ----------------------------------------------------------------
def equipes_em_viagem(
    viagens: Sequence[Viagem], banca: str, data: datetime.date
) -> dict[str, dict[tuple[str, int], int]]:
    """Examinadores da banca em viagem por período, agrupados por equipe.

    Uma equipe que atende dois municípios no mesmo dia em turnos diferentes é
    contada uma única vez (daí o ``max`` por equipe). A equipe é identificada
    por (banca titular, número), para não fundir a equipe 01 de Timon com a
    equipe 01 de Caxias em que Timon participa como apoio.

    A equipe ocupa os dois turnos em todos os dias do trecho, inclusive os de
    deslocamento e os de turno único: ela está fora da sede o dia inteiro.
    """
    resultado: dict[str, dict[tuple[str, int], int]] = {
        TURNO_MANHA: {},
        TURNO_TARDE: {},
    }
    for v in viagens_fora_da_sede(viagens, banca, data):
        quantidade = examinadores_da_banca_na_viagem(v, banca)
        equipe = chave_equipe(v)
        for periodo in (TURNO_MANHA, TURNO_TARDE):
            atual = resultado[periodo].get(equipe, 0)
            resultado[periodo][equipe] = max(atual, quantidade)
    return resultado


def efetivo_em_viagem(
    viagens: Sequence[Viagem], banca: str, data: datetime.date
) -> tuple[int, int]:
    equipes = equipes_em_viagem(viagens, banca, data)
    return sum(equipes[TURNO_MANHA].values()), sum(equipes[TURNO_TARDE].values())


def pico_diario(manha: int, tarde: int) -> int:
    """O efetivo necessário no dia é o maior dos turnos, nunca a soma."""
    return max(manha, tarde)


def agrupar_trechos_por_equipe(
    viagens: Sequence[Viagem], banca: str | None = None
) -> list[dict[str, Any]]:
    """Consolida os trechos de cada equipe itinerante em um único registro."""
    grupos: dict[tuple[str, int], list[Viagem]] = {}
    for v in viagens:
        banca_v = v.get("Banca", "")
        if banca and banca_v != banca:
            continue
        chave = (banca_v, int(v.get("Numero Banca Itinerante", 1) or 1))
        grupos.setdefault(chave, []).append(v)

    consolidados = []
    for (banca_v, numero), trechos in sorted(grupos.items()):
        trechos = sorted(trechos, key=lambda t: para_data(t["Data Inicio"]))
        inicio = min(para_data(t["Data Inicio"]) for t in trechos)
        fim = max(para_data(t["Data Fim"]) for t in trechos)
        destinos: list[str] = []
        for t in trechos:
            destino = t.get("Destino", "")
            if destino and destino not in destinos:
                destinos.append(destino)
        consolidados.append(
            {
                "Banca": banca_v,
                "Numero": numero,
                "Trechos": trechos,
                "Inicio": inicio,
                "Fim": fim,
                "Destinos": destinos,
                "Examinadores": max(
                    (total_examinadores(t) for t in trechos), default=0
                ),
            }
        )
    return consolidados


def equipes_ativas_no_intervalo(
    viagens: Sequence[Viagem],
    banca: str,
    inicio: datetime.date,
    fim: datetime.date,
) -> list[dict[str, Any]]:
    """Equipes itinerantes da banca que atuam dentro do intervalo informado."""
    ativas = []
    for grupo in agrupar_trechos_por_equipe(viagens):
        if not any(banca_vinculada(t, banca) for t in grupo["Trechos"]):
            continue
        dias = [
            d
            for t in grupo["Trechos"]
            for d in dias_efetivos_da_viagem(t)
            if inicio <= d <= fim
        ]
        if not dias:
            continue
        copia = dict(grupo)
        copia["Inicio Janela"] = min(dias)
        copia["Fim Janela"] = max(dias)
        copia["Examinadores"] = max(
            (
                examinadores_da_banca_na_viagem(t, banca)
                for t in grupo["Trechos"]
            ),
            default=0,
        )
        ativas.append(copia)
    return ativas


def formatar_datas_exames(dias_do_mes: Iterable[int]) -> str:
    """Agrupa [3,4,5,10] em '03 a 05; 10' para a publicação oficial."""
    dias = sorted({int(d) for d in dias_do_mes})
    if not dias:
        return "-"
    grupos: list[list[int]] = []
    atual = [dias[0]]
    for dia in dias[1:]:
        if dia == atual[-1] + 1:
            atual.append(dia)
        else:
            grupos.append(atual)
            atual = [dia]
    grupos.append(atual)

    partes = []
    for grupo in grupos:
        if len(grupo) == 1:
            partes.append(f"{grupo[0]:02d}")
        elif len(grupo) == 2:
            partes.append(f"{grupo[0]:02d} e {grupo[1]:02d}")
        else:
            partes.append(f"{grupo[0]:02d} a {grupo[-1]:02d}")
    return "; ".join(partes)


def bancas_para_validar(banca: str, bancas_apoio: Sequence[str] | None) -> list[str]:
    bancas = [banca]
    for outra in bancas_apoio or []:
        if outra not in bancas:
            bancas.append(outra)
    return bancas


def bancas_apoio_disponiveis(banca: str) -> list[str]:
    if banca not in BANCAS_APOIO_CONJUNTO:
        return []
    return [b for b in BANCAS_APOIO_CONJUNTO if b != banca]


# ===========================================================================
# 2.1 MOTOR DE CAPACIDADE E GERADOR AUTOMÁTICO DE CALENDÁRIO
# ===========================================================================
# Regras da especificação "Módulo de agendamento e gerador de calendário":
#   * Cat B, C, D, E: 1 examinador = 2 exames por horário.
#   * Cat A: 1 dupla = 12 exames por horário (1 examinador = 6), limitada pelo
#     número de pistas do local (1 pista comporta 1 dupla).
#   * Horários de fim de turno (11:30 e 16:30, configuráveis): capacidade de
#     B, C, D e E cai pela metade. Cat A não sofre o corte.
#   * A, C, D e E só de manhã. A tarde fica para a Cat B.
#   * C, D e E dividindo horário com a B saem com horário quebrado
#     (+1, +2 e +10 minutos).
#   * A soma da carga de todas as categorias no horário não pode passar do
#     número de examinadores do horário.
# A carga é medida em sextos de examinador (inteiros, sem erro de
# arredondamento): vaga A = 1; vaga B/C/D/E = 3; no fim de turno, 6.


def _hora_em_minutos(horario: str) -> int | None:
    try:
        hora, minuto = (int(p) for p in str(horario).strip().split(":")[:2])
    except (TypeError, ValueError):
        return None
    if not (0 <= hora <= 23 and 0 <= minuto <= 59):
        return None
    return hora * 60 + minuto


def _minutos_em_hora(minutos: int) -> str:
    minutos = max(0, min(int(minutos), 23 * 60 + 59))
    return f"{minutos // 60:02d}:{minutos % 60:02d}"


def categoria_da_coluna(coluna: str) -> str | None:
    """'Cat A' e 'PCD A' devolvem 'A'; qualquer outra coluna devolve None."""
    partes = str(coluna).split()
    if len(partes) == 2 and partes[0] in ("Cat", "PCD") and partes[1] in CATEGORIAS:
        return partes[1]
    return None


def eh_fim_de_turno(horario: str, horarios_fim_turno: Iterable[str]) -> bool:
    minutos = _hora_em_minutos(horario)
    return minutos is not None and minutos in {
        _hora_em_minutos(h) for h in horarios_fim_turno
    }


def sofre_corte(categoria: str, fim_turno: bool, corte_a: bool = False) -> bool:
    """A vaga cai pela metade neste horário? B a E sempre no fim de turno; a
    Cat A só nos postos com corte (Pátio do DETRAN às 11:30)."""
    if not fim_turno:
        return False
    return categoria in CATEGORIAS_COM_CORTE_FIM_TURNO or (categoria == "A" and corte_a)


def carga_por_vaga(categoria: str, fim_turno: bool, corte_a: bool = False) -> int:
    """Carga de uma vaga em sextos de examinador."""
    carga = UNIDADES_POR_EXAMINADOR // EXAMES_POR_EXAMINADOR[categoria]
    if sofre_corte(categoria, fim_turno, corte_a):
        carga *= 2
    return carga


def bloco_cat_a(fim_turno: bool, corte_a: bool = False) -> int:
    """Vagas de Cat A de uma pista (dupla) no horário: 12, ou 6 no fim de
    turno dos postos com corte."""
    return (
        VAGAS_CAT_A_POR_DUPLA // 2
        if sofre_corte("A", fim_turno, corte_a)
        else VAGAS_CAT_A_POR_DUPLA
    )


def capacidade_categoria(
    categoria: str,
    examinadores: int,
    fim_turno: bool,
    pistas_moto: int | None = None,
    corte_a: bool = False,
) -> int:
    """Vagas que ``examinadores`` conseguem atender sozinhos na categoria.

    Exemplos (6 examinadores): B = 12, B às 11:30 = 6, A = 36 (3 pistas).
    Cat A só conta duplas completas.
    """
    efetivo = max(0, int(examinadores))
    if categoria == "A":
        efetivo -= efetivo % EXAMINADORES_POR_DUPLA_MOTO
        if pistas_moto is not None:
            efetivo = min(efetivo, max(0, int(pistas_moto)) * EXAMINADORES_POR_DUPLA_MOTO)
    return efetivo * UNIDADES_POR_EXAMINADOR // carga_por_vaga(categoria, fim_turno, corte_a)


def teto_cat_a_por_pistas(
    pistas_moto: int | None, fim_turno: bool, corte_a: bool = False
) -> int | None:
    """Limite físico de vagas de Cat A no horário (None = sem limite)."""
    if pistas_moto is None:
        return None
    return capacidade_categoria(
        "A", int(pistas_moto) * EXAMINADORES_POR_DUPLA_MOTO, fim_turno, pistas_moto, corte_a
    )


def examinadores_do_horario(linha: dict[str, Any]) -> int:
    """Examinadores escalados no horário: Exam. M de manhã, Exam. T à tarde."""
    coluna = "Exam. M" if periodo_do_horario(linha.get("Horário")) == TURNO_MANHA else "Exam. T"
    return _inteiro(linha.get(coluna))


def vagas_por_categoria(linha: dict[str, Any]) -> dict[str, int]:
    """Soma Cat X + PCD X de cada categoria (PCD ocupa o examinador igual)."""
    totais = {c: 0 for c in CATEGORIAS}
    for coluna in COLS_VAGAS:
        categoria = categoria_da_coluna(coluna)
        if categoria:
            totais[categoria] += max(0, _inteiro(linha.get(coluna)))
    return totais


def consumo_em_unidades(
    linha: dict[str, Any], horarios_fim_turno: Iterable[str], corte_a: bool = False
) -> int:
    fim = eh_fim_de_turno(linha.get("Horário"), horarios_fim_turno)
    return sum(
        quantidade * carga_por_vaga(categoria, fim, corte_a)
        for categoria, quantidade in vagas_por_categoria(linha).items()
    )


def examinadores_necessarios(
    linha: dict[str, Any], horarios_fim_turno: Iterable[str], corte_a: bool = False
) -> float:
    return consumo_em_unidades(linha, horarios_fim_turno, corte_a) / UNIDADES_POR_EXAMINADOR


def validar_linha(
    linha: dict[str, Any],
    horarios_fim_turno: Iterable[str],
    pistas_moto: int | None = None,
    corte_a: bool = False,
) -> list[str]:
    """Problemas de capacidade da linha (vazio = linha válida).

    Linhas bloqueadas (feriado/indisponível) não ofertam vagas e por isso
    nunca acusam estouro, mesmo guardando os números lançados antes.
    """
    if linha.get("Status") != STATUS_DISPONIVEL:
        return []
    problemas: list[str] = []
    horario = str(linha.get("Horário", ""))
    fim = eh_fim_de_turno(horario, horarios_fim_turno)
    vagas = vagas_por_categoria(linha)
    examinadores = examinadores_do_horario(linha)

    usado = consumo_em_unidades(linha, horarios_fim_turno, corte_a)
    disponivel = examinadores * UNIDADES_POR_EXAMINADOR
    if usado > disponivel:
        problemas.append(
            "capacidade estourada: exige "
            + f"{usado / UNIDADES_POR_EXAMINADOR:.1f}".replace(".", ",")
            + f" examinador(es) e há {examinadores}"
        )

    if periodo_do_horario(horario) == TURNO_TARDE:
        indevidas = [c for c in sorted(CATEGORIAS_SO_MANHA) if vagas[c] > 0]
        if indevidas:
            problemas.append(
                "Cat " + ", ".join(indevidas) + " só pode ser agendada de manhã"
            )

    teto_a = teto_cat_a_por_pistas(pistas_moto, fim, corte_a)
    if teto_a is not None and vagas["A"] > teto_a:
        problemas.append(
            f"Cat A acima do limite físico de {pistas_moto} pista(s): máximo"
            f" {teto_a} por horário"
        )
    bloco = bloco_cat_a(fim, corte_a)
    if vagas["A"] % bloco:
        problemas.append(
            f"Cat A sai em blocos fechados de {bloco} por pista (dupla) neste"
            f" horário; {vagas['A']} não fecha o bloco"
        )
    return problemas


def horario_com_offset(horario: str, categoria: str) -> str:
    minutos = _hora_em_minutos(horario)
    if minutos is None:
        return str(horario)
    return _minutos_em_hora(minutos + OFFSET_MINUTOS.get(categoria, 0))


def horarios_quebrados(linha: dict[str, Any]) -> dict[str, str]:
    """Horário publicado de C, D e E quando dividem o horário com a Cat B."""
    vagas = vagas_por_categoria(linha)
    if vagas["B"] <= 0:
        return {}
    horario = str(linha.get("Horário", ""))
    return {
        c: horario_com_offset(horario, c)
        for c in CATEGORIAS
        if vagas[c] > 0 and OFFSET_MINUTOS.get(c, 0)
    }


def descricao_horarios_quebrados(linha: dict[str, Any]) -> str:
    return " · ".join(f"{c} {h}" for c, h in horarios_quebrados(linha).items())


def horarios_na_janela(horarios: Sequence[str], inicio: str, fim: str) -> list[str]:
    """Horários ativos dentro da janela de atendimento do dia (inclusiva)."""
    ini = _hora_em_minutos(inicio)
    ter = _hora_em_minutos(fim)
    if ini is None or ter is None:
        return list(horarios)
    return [h for h in horarios if (m := _hora_em_minutos(h)) is not None and ini <= m <= ter]


def _fator_turno(
    categoria: str,
    linhas_turno: Sequence[dict],
    fim_turno: Iterable[str],
    corte_a: bool = False,
) -> float:
    """Horários-equivalentes do turno para a categoria (corte vale 0,5)."""
    fim_turno = list(fim_turno)
    total = 0.0
    for linha in linhas_turno:
        fim = eh_fim_de_turno(linha.get("Horário"), fim_turno)
        total += 0.5 if sofre_corte(categoria, fim, corte_a) else 1.0
    return total


def _unidades_livres(
    linha: dict[str, Any], horarios_fim_turno: Iterable[str], corte_a: bool = False
) -> int:
    """Carga ainda livre no horário depois de tudo o que já está lançado nele
    (vagas PCD manuais e o que o gerador já colocou)."""
    usado = consumo_em_unidades(linha, horarios_fim_turno, corte_a)
    return max(0, examinadores_do_horario(linha) * UNIDADES_POR_EXAMINADOR - usado)


def _distribuir_duplas(total_duplas: int, quantidade_dias: int) -> list[int]:
    """Espalha ``total_duplas`` dias-dupla pelos dias de forma uniforme e
    centrada (7 em 22 dias: 2, 5, 8, 11, 15, 18 e 21), sem amontoar no
    começo nem no fim do período."""
    if quantidade_dias <= 0:
        return []
    acumulado = [
        ((i + 1) * 2 * total_duplas + quantidade_dias) // (2 * quantidade_dias)
        for i in range(quantidade_dias)
    ]
    return [acumulado[0]] + [acumulado[i] - acumulado[i - 1] for i in range(1, quantidade_dias)]


def gerar_distribuicao(
    linhas: Sequence[dict[str, Any]],
    demanda_mensal: dict[str, int],
    horarios_fim_turno: Iterable[str],
    pistas_moto: int | None = None,
    *,
    periodo: tuple[datetime.date, datetime.date] | None = None,
    dias_por_categoria: dict[str, Iterable[str]] | None = None,
    corte_a: bool = False,
) -> tuple[list[dict[str, Any]], dict[str, int], dict[str, int]]:
    """Distribui a demanda mensal de cada categoria pelos horários disponíveis.

    ``demanda_mensal`` tem as chaves "A" a "E" e, opcionalmente, "PCD B".
    ``periodo`` (início, fim) limita os dias gerados; linhas fora dele não são
    tocadas, o que permite gerar o mês em partes. ``dias_por_categoria``
    restringe cada categoria a alguns dias da semana (ex.: Cat A só terça e
    quinta, PCD B só quinta). ``corte_a`` = a Cat A também cai pela metade no
    fim de turno (Pátio do DETRAN).

    Algoritmo (nivelado, dia a dia):
      1. Cat A sai em blocos fechados: cada pista (dupla) oferece 12 vagas em
         todos os horários da manhã do dia (6 no fim de turno com corte). Os
         dias com Cat A são espalhados pelo período e o último bloco fecha
         inteiro, então o total pode passar da demanda em menos de um bloco.
      2. As demais categorias recebem por dia ``ceil(restante / dias que
         restam para a categoria)``. O que um dia não atende passa adiante.
      3. Tarde: só Cat B, até a capacidade de cada horário.
      4. Manhã: PCD B primeiro; depois os examinadores livres são divididos
         entre C, D, E e o que sobrar da meta da B, em quantidade fixa para
         todos os horários da manhã. Se faltar gente, a B cede primeiro (ela
         tem a tarde) e depois a categoria que mais consome.
      5. Cada horário recebe a sua cota da meta do dia (proporcional ao peso
         dos horários que faltam no turno), limitada à capacidade dele.

    Só mexe nas colunas Cat A..E (e PCD B, quando pedida) de linhas
    Disponível dentro do período. As outras colunas PCD continuam manuais e a
    carga delas é descontada antes.
    Devolve (linhas, vagas geradas por chave, déficit por chave).
    """
    fim_turno = list(horarios_fim_turno)
    resultado = [dict(linha) for linha in linhas]
    gera_pcd = COL_PCD_GERADA in demanda_mensal
    chaves = CATEGORIAS + ([COL_PCD_GERADA] if gera_pcd else [])
    restante = {c: max(0, _inteiro(demanda_mensal.get(c))) for c in chaves}
    gerado = {c: 0 for c in chaves}
    filtro_dias = {
        c: set(dias) for c, dias in (dias_por_categoria or {}).items() if dias is not None
    }

    por_dia: dict[datetime.date, list[dict[str, Any]]] = {}
    for linha in resultado:
        if linha.get("Status") != STATUS_DISPONIVEL:
            continue
        data = para_data(linha.get("Data"))
        if data is None or (periodo and not (periodo[0] <= data <= periodo[1])):
            continue
        for c in CATEGORIAS:
            linha[f"Cat {c}"] = 0
        if gera_pcd:
            linha[COL_PCD_GERADA] = 0
        if examinadores_do_horario(linha) > 0:
            por_dia.setdefault(data, []).append(linha)

    def ordenadas(dia: datetime.date, periodo_turno: str) -> list[dict[str, Any]]:
        return sorted(
            (l for l in por_dia[dia] if periodo_do_horario(l["Horário"]) == periodo_turno),
            key=lambda l: _hora_em_minutos(l["Horário"]) or 0,
        )

    def so_manha(chave: str) -> bool:
        return chave == COL_PCD_GERADA or chave in CATEGORIAS_SO_MANHA

    dias = sorted(por_dia)
    elegiveis = {
        c: [
            d
            for d in dias
            if (c not in filtro_dias or nome_dia_semana(d) in filtro_dias[c])
            and (not so_manha(c) or ordenadas(d, TURNO_MANHA))
        ]
        for c in chaves
    }

    def bloco(linha: dict[str, Any]) -> int:
        return bloco_cat_a(eh_fim_de_turno(linha["Horário"], fim_turno), corte_a)

    # Plano da Cat A: quantas duplas trabalham em cada dia elegível.
    dias_a = elegiveis["A"]
    capacidade_dupla_dia = {d: sum(bloco(l) for l in ordenadas(d, TURNO_MANHA)) for d in dias_a}
    plano_a: dict[datetime.date, int] = {}
    if restante["A"] and dias_a:
        media = sum(capacidade_dupla_dia.values()) / len(dias_a)
        total_duplas = math.ceil(restante["A"] / media) if media else 0
        plano_a = dict(zip(dias_a, _distribuir_duplas(total_duplas, len(dias_a))))

    for dia in dias:
        manha = ordenadas(dia, TURNO_MANHA)
        tarde = ordenadas(dia, TURNO_TARDE)
        meta = {}
        for c in chaves:
            if dia in elegiveis[c]:
                faltam = sum(1 for d in elegiveis[c] if d >= dia)
                meta[c] = math.ceil(restante[c] / faltam)
            else:
                meta[c] = 0

        def lancar(linha: dict[str, Any], chave: str, quantidade: int) -> None:
            quantidade = max(0, min(quantidade, meta[chave], restante[chave]))
            if quantidade <= 0:
                return
            coluna = chave if chave == COL_PCD_GERADA else f"Cat {chave}"
            linha[coluna] = _inteiro(linha.get(coluna)) + quantidade
            meta[chave] -= quantidade
            restante[chave] -= quantidade
            gerado[chave] += quantidade

        def cota(chave: str, linhas_turno: Sequence[dict], posicao: int) -> int:
            """Parte da meta que cabe a este horário, proporcional ao peso dos
            horários que faltam no turno (fim de turno pesa 0,5 em B a E).
            Espalha as vagas pelo turno em vez de lotar o primeiro horário."""
            categoria = "B" if chave == COL_PCD_GERADA else chave
            restantes = linhas_turno[posicao:]
            peso_total = _fator_turno(categoria, restantes, fim_turno, corte_a)
            peso = _fator_turno(categoria, restantes[:1], fim_turno, corte_a)
            if peso_total <= 0:
                return 0
            return math.ceil(meta[chave] * peso / peso_total)

        # 1. Cat A em blocos fechados de dupla.
        if manha and plano_a.get(dia) is not None and restante["A"] > 0:
            futuro = sum(
                plano_a[d] * capacidade_dupla_dia[d] for d in dias_a if d > dia
            )
            duplas = plano_a[dia]
            if restante["A"] > duplas * capacidade_dupla_dia[dia] + futuro:
                duplas = math.ceil((restante["A"] - futuro) / capacidade_dupla_dia[dia])
            livres = min(_unidades_livres(l, fim_turno, corte_a) for l in manha)
            maximo = livres // UNIDADES_POR_EXAMINADOR // EXAMINADORES_POR_DUPLA_MOTO
            if pistas_moto is not None:
                maximo = min(maximo, max(0, int(pistas_moto)))
            duplas = max(0, min(duplas, maximo))
            for linha in manha:
                if duplas <= 0 or restante["A"] <= 0:
                    break
                tamanho = bloco(linha)
                quantidade = min(duplas, math.ceil(restante["A"] / tamanho)) * tamanho
                linha["Cat A"] = _inteiro(linha.get("Cat A")) + quantidade
                restante["A"] -= quantidade
                gerado["A"] += quantidade

        # 2. Tarde: só Cat B.
        for posicao, linha in enumerate(tarde):
            fim = eh_fim_de_turno(linha["Horário"], fim_turno)
            capacidade = _unidades_livres(linha, fim_turno, corte_a) // carga_por_vaga("B", fim)
            lancar(linha, "B", min(capacidade, cota("B", tarde, posicao)))

        if not manha:
            continue

        # 3. PCD B (manhã), antes da ampla concorrência.
        if gera_pcd:
            for posicao, linha in enumerate(manha):
                fim = eh_fim_de_turno(linha["Horário"], fim_turno)
                capacidade = _unidades_livres(linha, fim_turno, corte_a) // carga_por_vaga("B", fim)
                lancar(linha, COL_PCD_GERADA, min(capacidade, cota(COL_PCD_GERADA, manha, posicao)))

        # 4. C, D, E e B com examinadores fixos para toda a manhã.
        livres = min(
            _unidades_livres(l, fim_turno, corte_a) // UNIDADES_POR_EXAMINADOR for l in manha
        )
        ordem = ["C", "D", "E", "B"]
        alocacao: dict[str, int] = {}
        for c in ordem:
            if meta[c] <= 0:
                alocacao[c] = 0
                continue
            por_examinador = EXAMES_POR_EXAMINADOR[c] * _fator_turno(c, manha, fim_turno, corte_a)
            alocacao[c] = math.ceil(meta[c] / por_examinador) if por_examinador else 0

        while sum(alocacao.values()) > livres:
            if alocacao["B"] > 0:
                alocacao["B"] -= 1
                continue
            maior = max((c for c in ordem if alocacao[c] > 0), key=lambda c: alocacao[c])
            alocacao[maior] -= 1

        for posicao, linha in enumerate(manha):
            fim = eh_fim_de_turno(linha["Horário"], fim_turno)
            for c in ordem:
                if alocacao[c] > 0:
                    lancar(
                        linha,
                        c,
                        min(
                            capacidade_categoria(c, alocacao[c], fim, None, corte_a),
                            cota(c, manha, posicao),
                        ),
                    )

    for linha in resultado:
        if linha.get("Status") == STATUS_DISPONIVEL:
            linha["Total"] = sum(max(0, _inteiro(linha.get(c))) for c in COLS_VAGAS)
    return resultado, gerado, {c: max(0, v) for c, v in restante.items()}


# --- Efetivo diário: fonte única para tela, quadro e PDF -------------------
# Antes, cada tela contava o efetivo de um jeito (com ou sem linhas vazias,
# com ou sem dias fora da grade) e os números não batiam. Regra única agora:
#   * Só contam linhas Disponível com vagas lançadas (Total > 0). Uma
#     localidade só aberta, sem vagas, não ocupa examinador.
#   * Manhã = maior Exam. M entre os horários da manhã com vagas; tarde, idem.
#   * Localidade atendida por equipe itinerante não soma pelo município: a
#     equipe entra uma vez só (ver ``equipes_em_viagem``).
#   * Os dados devem chegar já filtrados pela grade vigente (sem linhas de
#     horários desativados ou dias removidos).


def efetivo_do_grupo(grupo_do_dia: pd.DataFrame | Sequence[dict]) -> tuple[int, int]:
    """(manhã, tarde) de uma localidade em um dia (DataFrame ou registros)."""
    if grupo_do_dia is None or len(grupo_do_dia) == 0:
        return 0, 0
    registros = (
        grupo_do_dia.to_dict(orient="records")
        if isinstance(grupo_do_dia, pd.DataFrame)
        else grupo_do_dia
    )
    manha = tarde = 0
    for registro in registros:
        if registro.get("Status") != STATUS_DISPONIVEL or _inteiro(registro.get("Total")) <= 0:
            continue
        if periodo_do_horario(registro.get("Horário")) == TURNO_MANHA:
            manha = max(manha, _inteiro(registro.get("Exam. M")))
        else:
            tarde = max(tarde, _inteiro(registro.get("Exam. T")))
    return manha, tarde


def registros_por_data(df: pd.DataFrame | None) -> dict[str, list[dict]]:
    """Agrupa as linhas por data (texto dd/mm/aaaa), sem custo de pandas."""
    agrupado: dict[str, list[dict]] = {}
    if df is None or df.empty:
        return agrupado
    for registro in df.to_dict(orient="records"):
        agrupado.setdefault(str(registro.get("Data")), []).append(registro)
    return agrupado


def efetivo_diario(
    dados_por_local: dict[str, pd.DataFrame],
    viagens: Sequence[Viagem],
    banca: str,
    ano: int,
    mes_numero: int,
    locais_fixos: Iterable[str] | None = None,
) -> dict[datetime.date, dict[str, Any]]:
    """Efetivo da banca por dia do mês.

    Devolve {data: {M, T, locais, equipes, viagem, sem_viagem}}:
      * ``locais``: (manhã, tarde) de cada localidade que somou pelo próprio
        calendário;
      * ``equipes``: (manhã, tarde) de cada equipe itinerante fora da sede;
      * ``viagem``: examinadores fora da sede no dia (equipes + municípios
        itinerantes com vagas sem viagem cadastrada), a linha "Viagens" do
        quadro;
      * ``sem_viagem``: municípios itinerantes que tiveram vagas sem viagem.
    ``locais_fixos`` (opcional) separa postos fixos de municípios
    itinerantes; sem ele, toda localidade conta como posto fixo.
    """
    fixos = set(locais_fixos) if locais_fixos is not None else None
    totais: dict[datetime.date, dict[str, Any]] = {}

    def acumulado(data: datetime.date) -> dict[str, Any]:
        return totais.setdefault(
            data,
            {"M": 0, "T": 0, "locais": {}, "equipes": {}, "viagem": 0, "sem_viagem": []},
        )

    for local, df in dados_por_local.items():
        viagens_do_local = [
            v for v in viagens if banca_vinculada(v, banca) and v.get("Destino") == local
        ]
        for data_texto, grupo in registros_por_data(df).items():
            data = para_data(data_texto)
            if not data or data.month != mes_numero or data.year != int(ano):
                continue
            # Localidade com equipe itinerante no dia conta pela equipe.
            if any(viagem_presente_no_dia(v, data) for v in viagens_do_local):
                continue
            manha, tarde = efetivo_do_grupo(grupo)
            if manha or tarde:
                ac = acumulado(data)
                ac["M"] += manha
                ac["T"] += tarde
                ac["locais"][local] = (manha, tarde)
                if fixos is not None and local not in fixos:
                    ac["viagem"] += pico_diario(manha, tarde)
                    ac["sem_viagem"].append(local)

    for numero_dia in range(1, calendar.monthrange(int(ano), mes_numero)[1] + 1):
        data = datetime.date(int(ano), mes_numero, numero_dia)
        equipes = equipes_em_viagem(viagens, banca, data)
        manha = sum(equipes[TURNO_MANHA].values())
        tarde = sum(equipes[TURNO_TARDE].values())
        if manha or tarde:
            ac = acumulado(data)
            ac["M"] += manha
            ac["T"] += tarde
            ac["viagem"] += pico_diario(manha, tarde)
            for chave in set(equipes[TURNO_MANHA]) | set(equipes[TURNO_TARDE]):
                ac["equipes"][chave] = (
                    equipes[TURNO_MANHA].get(chave, 0),
                    equipes[TURNO_TARDE].get(chave, 0),
                )
    return totais


def pico_do_mes(totais: dict[datetime.date, dict[str, Any]]) -> int:
    if not totais:
        return 0
    return max(pico_diario(v["M"], v["T"]) for v in totais.values())


# --- Ordenação para os relatórios em PDF ------------------------------------
def _primeira_data_lancada(df: pd.DataFrame) -> datetime.date:
    """Data do primeiro exame disponível da localidade (para ordenar)."""
    if df is None or df.empty:
        return datetime.date.max
    disponiveis = df[(df["Status"] == STATUS_DISPONIVEL) & (df["Total"] > 0)]
    if disponiveis.empty:
        return datetime.date.max
    datas = [para_data(d) for d in disponiveis["Data"].unique()]
    validas = [d for d in datas if d]
    return min(validas) if validas else datetime.date.max


def ordenar_localidades_para_pdf(
    banca: str,
    bancas_config: dict[str, list[str]],
    dados_por_local: dict[str, pd.DataFrame],
) -> list[tuple[str, str]]:
    """Ordem de apresentação das localidades no PDF da banca.

    Regra: unidades da sede primeiro, depois a região metropolitana e, por
    fim, os demais municípios pela data do primeiro exame lançado.
    Devolve pares (grupo, localidade).
    """
    prioridade = ORDEM_PRIORITARIA_PDF.get(banca)
    if not prioridade:
        sede = obter_sede(bancas_config, banca)
        prioridade = [sede] if sede else []

    indice_prioridade = {normalizar_texto(nome): i for i, nome in enumerate(prioridade)}

    prioritarias: list[tuple[int, str]] = []
    demais: list[tuple[datetime.date, str, str]] = []

    for local, df in dados_por_local.items():
        chave = normalizar_texto(local)
        if chave in indice_prioridade:
            prioritarias.append((indice_prioridade[chave], local))
        else:
            demais.append((_primeira_data_lancada(df), normalizar_texto(local), local))

    ordenadas: list[tuple[str, str]] = []
    for _, local in sorted(prioritarias):
        grupo = GRUPOS_PDF.get(local)
        if grupo is None:
            grupo = GRUPO_SEDE_GENERICO
        ordenadas.append((grupo, local))
    for _, _, local in sorted(demais):
        ordenadas.append((GRUPO_INTERIOR, local))
    return ordenadas


# --- Planilha oficial de publicação -----------------------------------------
def chave_ordem_publicacao(
    banca: str, local: str, bancas_config: dict[str, list[str]]
) -> tuple:
    """Ordem da planilha: São Luís, Imperatriz, Timon, Bacabal, Caxias e Santa
    Inês. Em São Luís, os sete postos da região metropolitana na ordem
    oficial; nas demais, o município da banca primeiro. O resto em ordem
    alfabética."""
    try:
        posicao_banca = ORDEM_BANCAS_PUBLICACAO.index(banca)
    except ValueError:
        posicao_banca = len(ORDEM_BANCAS_PUBLICACAO)
    prioridade = list(ORDEM_PRIORITARIA_PUBLICACAO.get(banca, []))
    if not prioridade:
        prioridade = [
            nome
            for nome in ((BANCAS_PADRAO.get(banca) or [None])[0], obter_sede(bancas_config, banca))
            if nome
        ]
    indices = {normalizar_texto(nome): i for i, nome in enumerate(prioridade)}
    posicao_local = indices.get(normalizar_texto(local), len(prioridade))
    return (posicao_banca, normalizar_texto(banca), posicao_local, normalizar_texto(local))


_EXTENSO_UNIDADES = [
    "zero", "uma", "duas", "três", "quatro", "cinco", "seis", "sete", "oito",
    "nove", "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis",
    "dezessete", "dezoito", "dezenove",
]
_EXTENSO_DEZENAS = [
    "", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta",
    "oitenta", "noventa",
]
_EXTENSO_CENTENAS = [
    "", "cento", "duzentas", "trezentas", "quatrocentas", "quinhentas",
    "seiscentas", "setecentas", "oitocentas", "novecentas",
]


def _extenso_ate_999(numero: int) -> str:
    if numero == 100:
        return "cem"
    centena, resto = divmod(numero, 100)
    partes = [_EXTENSO_CENTENAS[centena]] if centena else []
    if resto:
        if resto < 20:
            partes.append(_EXTENSO_UNIDADES[resto])
        else:
            dezena, unidade = divmod(resto, 10)
            partes.append(
                _EXTENSO_DEZENAS[dezena]
                + (f" e {_EXTENSO_UNIDADES[unidade]}" if unidade else "")
            )
    return " e ".join(partes)


def numero_por_extenso(numero: int) -> str:
    """Número por extenso no feminino (concorda com "vagas"): 32 → "trinta e
    duas", 1200 → "mil e duzentas". Vai até 999.999."""
    numero = int(numero)
    if numero <= 0:
        return "zero"
    milhares, resto = divmod(numero, 1000)
    partes = []
    if milhares:
        partes.append("mil" if milhares == 1 else f"{_extenso_ate_999(milhares)} mil")
    if resto:
        texto = _extenso_ate_999(resto)
        if milhares and (resto < 100 or resto % 100 == 0):
            texto = "e " + texto
        partes.append(texto)
    return " ".join(partes)


DIA_POR_EXTENSO = {
    "Segunda": "segunda-feira",
    "Terça": "terça-feira",
    "Quarta": "quarta-feira",
    "Quinta": "quinta-feira",
    "Sexta": "sexta-feira",
    "Sábado": "sábado",
    "Domingo": "domingo",
}


def _juntar_com_e(itens: Sequence[str]) -> str:
    itens = list(itens)
    if len(itens) <= 1:
        return "".join(itens)
    return ", ".join(itens[:-1]) + " e " + itens[-1]


def nome_publicacao_local(local: str) -> str:
    """Como a localidade aparece no texto da observação."""
    if normalizar_texto(local) == normalizar_texto(LOCAL_SL_PATIO):
        return "no Pátio do DETRAN/MA"
    return f"em {local}"


def observacao_pcd(
    marcador: str,
    categoria: str,
    local: str,
    quantidade: int,
    dias_semana: Iterable[str],
    turnos: Iterable[str],
) -> str:
    """Legenda do asterisco da planilha de publicação.

    No caso padrão (Cat B no Pátio, quintas pela manhã) sai exatamente o texto
    oficial; muda só a quantidade. Outras categorias, postos, dias ou turnos
    seguem o mesmo modelo.
    """
    dias = [DIA_POR_EXTENSO[d] for d in DIAS_SEMANA_OPCOES if d in set(dias_semana)]
    turnos = set(turnos)
    if turnos == {TURNO_MANHA}:
        periodo = "nas manhãs de "
    elif turnos == {TURNO_TARDE}:
        periodo = "nas tardes de "
    else:
        periodo = "nas manhãs e tardes de "
    quando = periodo + _juntar_com_e(dias) if dias else "conforme o calendário do posto"
    numero = f"{int(quantidade):,}".replace(",", ".")
    extenso = numero_por_extenso(quantidade)
    inicio = (
        f"{marcador} Observação: Do quantitativo total de vagas ofertadas para a"
        f" Categoria {categoria} {nome_publicacao_local(local)}, "
    )
    if int(quantidade) == 1:
        return (
            inicio
            + f"{numero} ({extenso}) vaga é destinada exclusivamente a candidatos"
            f" Pessoa com Deficiência (PcD), sendo disponibilizada {quando}."
        )
    return (
        inicio
        + f"{numero} ({extenso}) vagas são destinadas exclusivamente a candidatos"
        " Pessoa com Deficiência (PcD), sendo distribuídas ao longo do mês e"
        f" disponibilizadas {quando}."
    )


# ===========================================================================
# 3. PERSISTÊNCIA
# ===========================================================================
# SQLite com gravação por entidade e só quando o conteúdo muda de fato.

MAX_SNAPSHOTS = 20
_LOCK = threading.Lock()


class BackupInvalido(Exception):
    """Arquivo de backup fora do formato esperado."""


def caminho_banco() -> Path:
    return Path(os.environ.get("DETRAN_DB", "dados/detran.db"))


def pasta_snapshots() -> Path:
    return Path(os.environ.get("DETRAN_SNAPSHOTS", "dados/snapshots"))


ESQUEMA = """
CREATE TABLE IF NOT EXISTS config (
    chave TEXT PRIMARY KEY,
    valor TEXT NOT NULL,
    atualizado_em TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS bancas (
    nome  TEXT PRIMARY KEY,
    ordem INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS localidades (
    banca TEXT NOT NULL,
    nome  TEXT NOT NULL,
    ordem INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (banca, nome)
);

CREATE TABLE IF NOT EXISTS viagens (
    id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    banca                  TEXT NOT NULL,
    numero_equipe          INTEGER NOT NULL DEFAULT 1,
    destino                TEXT NOT NULL,
    data_inicio            TEXT NOT NULL,
    data_fim               TEXT NOT NULL,
    turno                  TEXT NOT NULL DEFAULT 'Dia inteiro',
    turnos_por_data        TEXT NOT NULL DEFAULT '{}',
    examinadores           INTEGER NOT NULL DEFAULT 0,
    bancas_apoio           TEXT NOT NULL DEFAULT '[]',
    examinadores_por_banca TEXT NOT NULL DEFAULT '{}',
    observacoes            TEXT NOT NULL DEFAULT '',
    horario_inicio         TEXT NOT NULL DEFAULT '',
    horario_fim            TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS historico (
    chave      TEXT NOT NULL,
    data       TEXT NOT NULL,
    horario    TEXT NOT NULL,
    dia_semana TEXT NOT NULL DEFAULT '',
    status     TEXT NOT NULL,
    exam_m     INTEGER NOT NULL DEFAULT 0,
    exam_t     INTEGER NOT NULL DEFAULT 0,
    vagas      TEXT NOT NULL DEFAULT '{}',
    total      INTEGER NOT NULL DEFAULT 0,
    bloqueio_auto INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (chave, data, horario)
);

CREATE TABLE IF NOT EXISTS parametros_local (
    chave TEXT PRIMARY KEY,
    valor TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS demandas (
    chave TEXT PRIMARY KEY,
    valor TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS dias_permitidos (
    chave TEXT PRIMARY KEY,
    dias  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS feriados_locais (
    chave TEXT PRIMARY KEY,
    texto TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS disponibilidade_semanal (
    chave TEXT PRIMARY KEY,
    valor INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS efetivo_diario (
    chave TEXT PRIMARY KEY,
    valor INTEGER NOT NULL
);
"""


def _migrar_esquema(conexao: sqlite3.Connection) -> None:
    """Adiciona colunas novas em bancos criados por versões anteriores."""
    colunas = {
        linha["name"] for linha in conexao.execute("PRAGMA table_info(viagens)")
    }
    if colunas and "turnos_por_data" not in colunas:
        LOG.info("Migrando esquema: adicionando viagens.turnos_por_data")
        with conexao:
            conexao.execute(
                "ALTER TABLE viagens ADD COLUMN turnos_por_data TEXT NOT NULL"
                " DEFAULT '{}'"
            )
    for coluna in ("horario_inicio", "horario_fim"):
        if colunas and coluna not in colunas:
            LOG.info("Migrando esquema: adicionando viagens.%s", coluna)
            with conexao:
                conexao.execute(
                    f"ALTER TABLE viagens ADD COLUMN {coluna} TEXT NOT NULL DEFAULT ''"
                )
    colunas_hist = {
        linha["name"] for linha in conexao.execute("PRAGMA table_info(historico)")
    }
    if colunas_hist and "bloqueio_auto" not in colunas_hist:
        # Linhas antigas marcadas como feriado eram sempre geradas pelo
        # sistema: passam a ser bloqueio automático, que se desfaz sozinho se
        # o feriado deixar de existir. "Indisponível" antigo pode ter sido
        # escolha do operador e fica como manual.
        LOG.info("Migrando esquema: adicionando historico.bloqueio_auto")
        with conexao:
            conexao.execute(
                "ALTER TABLE historico ADD COLUMN bloqueio_auto INTEGER NOT NULL"
                " DEFAULT 0"
            )
            conexao.execute(
                "UPDATE historico SET bloqueio_auto = 1 WHERE status = ?",
                (STATUS_FERIADO,),
            )


@st.cache_resource(show_spinner=False)
def _abrir_conexao(caminho: str) -> sqlite3.Connection:
    Path(caminho).parent.mkdir(parents=True, exist_ok=True)
    conexao = sqlite3.connect(caminho, check_same_thread=False)
    conexao.row_factory = sqlite3.Row
    conexao.execute("PRAGMA journal_mode=WAL")
    conexao.execute("PRAGMA synchronous=NORMAL")
    with conexao:
        conexao.executescript(ESQUEMA)
    _migrar_esquema(conexao)
    LOG.info("Banco aberto em %s", Path(caminho).resolve())
    return conexao


# Bancos cujo esquema já foi conferido por esta carga do código. A conexão
# fica em cache entre recargas do app; sem isto, atualizar o sistema com o
# servidor no ar deixaria a conexão antiga sem as migrações novas.
_ESQUEMA_CONFERIDO: set[str] = set()


def conectar() -> sqlite3.Connection:
    caminho = str(caminho_banco())
    conexao = _abrir_conexao(caminho)
    if caminho not in _ESQUEMA_CONFERIDO:
        with _LOCK:
            with conexao:
                conexao.executescript(ESQUEMA)
            _migrar_esquema(conexao)
        _ESQUEMA_CONFERIDO.add(caminho)
    return conexao


# --- Utilitários ------------------------------------------------------------
def _json_seguro(valor: Any) -> Any:
    if isinstance(valor, (datetime.date, datetime.datetime)):
        return valor.isoformat()
    if hasattr(valor, "item"):
        return valor.item()
    if isinstance(valor, (set, frozenset)):
        return sorted(valor)
    return str(valor)


def _dump(valor: Any) -> str:
    return json.dumps(valor, ensure_ascii=False, default=_json_seguro, allow_nan=False)


def _inteiro(valor: Any, padrao: int = 0) -> int:
    try:
        if valor is None or (isinstance(valor, float) and pd.isna(valor)):
            return padrao
        return int(valor)
    except (TypeError, ValueError):
        return padrao


def _assinatura(valor: Any) -> str:
    try:
        bruto = _dump(valor)
    except (TypeError, ValueError):
        bruto = repr(valor)
    return hashlib.sha1(bruto.encode("utf-8")).hexdigest()


def _mudou(nome: str, valor: Any) -> bool:
    """Só verifica se o conteúdo mudou — não marca nada como salvo.

    A confirmação de que os dados foram de fato persistidos é feita à
    parte, por `_confirmar_salvo`, e só deve acontecer depois que a
    gravação no banco tiver sucesso. Antes, `_mudou` já marcava a
    assinatura como "salva" antes mesmo de tentar gravar: se a gravação
    falhasse (disco cheio, banco bloqueado por outro processo etc.), o
    sistema passava a acreditar que aquele conteúdo já estava seguro no
    banco, e uma nova tentativa com o mesmo conteúdo era silenciosamente
    pulada — sem avisar que os dados nunca chegaram a ser salvos de verdade.
    """
    assinaturas = st.session_state.setdefault("_assinaturas", {})
    return assinaturas.get(nome) != _assinatura(valor)


def _confirmar_salvo(nome: str, valor: Any) -> None:
    """Registra que este conteúdo foi gravado com sucesso no banco."""
    st.session_state.setdefault("_assinaturas", {})[nome] = _assinatura(valor)


def _registrar_erro(mensagem: str, excecao: Exception) -> None:
    LOG.exception(mensagem)
    st.session_state["erro_persistencia"] = f"{mensagem}: {excecao}"


def _marcar_sucesso() -> None:
    st.session_state["erro_persistencia"] = None
    st.session_state["ultimo_salvamento"] = datetime.datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )


# --- Leitura ----------------------------------------------------------------
def _ler_config(conexao: sqlite3.Connection, chave: str, padrao: Any) -> Any:
    linha = conexao.execute(
        "SELECT valor FROM config WHERE chave = ?", (chave,)
    ).fetchone()
    if not linha:
        return padrao
    try:
        return json.loads(linha["valor"])
    except json.JSONDecodeError:
        LOG.warning("Config %s corrompida — usando padrão", chave)
        return padrao


def _ler_bancas(conexao: sqlite3.Connection) -> dict[str, list[str]]:
    bancas: dict[str, list[str]] = {}
    for linha in conexao.execute("SELECT nome FROM bancas ORDER BY ordem, nome"):
        bancas[linha["nome"]] = []
    for linha in conexao.execute(
        "SELECT banca, nome FROM localidades ORDER BY banca, ordem, nome"
    ):
        bancas.setdefault(linha["banca"], []).append(linha["nome"])
    return bancas


def _ler_viagens(conexao: sqlite3.Connection) -> list[Viagem]:
    viagens = []
    for linha in conexao.execute("SELECT * FROM viagens ORDER BY data_inicio, id"):
        try:
            bancas_apoio = json.loads(linha["bancas_apoio"])
            por_banca = json.loads(linha["examinadores_por_banca"])
        except json.JSONDecodeError:
            bancas_apoio, por_banca = [], {}
        try:
            turnos = json.loads(linha["turnos_por_data"])
        except (json.JSONDecodeError, IndexError, KeyError):
            turnos = {}
        colunas = linha.keys()
        viagens.append(
            {
                "id": linha["id"],
                "Banca": linha["banca"],
                "Numero Banca Itinerante": linha["numero_equipe"],
                "Destino": linha["destino"],
                "Data Inicio": para_data(linha["data_inicio"]),
                "Data Fim": para_data(linha["data_fim"]),
                "Turno": linha["turno"],
                "Turnos Por Data": turnos,
                "Examinadores": linha["examinadores"],
                "Bancas Apoio": bancas_apoio,
                "Examinadores Por Banca": por_banca,
                "Observações": linha["observacoes"],
                "Horario Inicio": linha["horario_inicio"] if "horario_inicio" in colunas else "",
                "Horario Fim": linha["horario_fim"] if "horario_fim" in colunas else "",
            }
        )
    return viagens


def _ler_historico(conexao: sqlite3.Connection) -> dict[str, pd.DataFrame]:
    por_chave: dict[str, list[dict[str, Any]]] = {}
    for linha in conexao.execute(
        "SELECT * FROM historico ORDER BY chave, data, horario"
    ):
        try:
            vagas = json.loads(linha["vagas"])
        except json.JSONDecodeError:
            vagas = {}
        registro = {
            "Data": linha["data"],
            "Dia da Semana": linha["dia_semana"],
            "Status": linha["status"],
            "Exam. M": linha["exam_m"],
            "Exam. T": linha["exam_t"],
            "Horário": linha["horario"],
        }
        for coluna in COLS_VAGAS:
            registro[coluna] = _inteiro(vagas.get(coluna, 0))
        registro["Total"] = linha["total"]
        registro[COL_AUTO] = _inteiro(linha["bloqueio_auto"])
        por_chave.setdefault(linha["chave"], []).append(registro)
    return {chave: pd.DataFrame(linhas) for chave, linhas in por_chave.items()}


# Tabelas chave/valor. Os nomes vêm só desta lista (nunca do usuário), o que
# torna seguro montá-los no SQL.
MAPAS = {
    # nome lógico: (tabela, coluna, valor é JSON?)
    "dias_permitidos": ("dias_permitidos", "dias", True),
    "feriados_locais": ("feriados_locais", "texto", False),
    "disponibilidade": ("disponibilidade_semanal", "valor", False),
    "parametros_local": ("parametros_local", "valor", True),
    "demandas": ("demandas", "valor", True),
    "efetivo_dia": ("efetivo_diario", "valor", False),
}


def _ler_mapa(conexao: sqlite3.Connection, nome: str) -> dict[str, Any]:
    tabela, coluna, eh_json = MAPAS[nome]
    resultado = {}
    for linha in conexao.execute(f"SELECT chave, {coluna} FROM {tabela}"):
        valor = linha[coluna]
        if eh_json:
            try:
                valor = json.loads(valor)
            except json.JSONDecodeError:
                LOG.warning("Valor corrompido em %s[%s]", tabela, linha["chave"])
                continue
        resultado[linha["chave"]] = valor
    return resultado


# --- Escrita ----------------------------------------------------------------
# Toda gravação é diferencial: compara com o que esta sessão leu ou gravou por
# último e mexe só nas linhas que ela mudou. Antes, viagens, bancas e mapas
# eram gravados com "apaga a tabela inteira e regrava a cópia da sessão", o
# que fazia uma sessão apagar o trabalho de outra que estivesse aberta ao
# mesmo tempo.


def salvar_config(chave: str, valor: Any, forcar: bool = False) -> None:
    if not forcar and not _mudou(f"config:{chave}", valor):
        return
    try:
        conexao = conectar()
        with _LOCK, conexao:
            conexao.execute(
                "INSERT INTO config (chave, valor, atualizado_em) VALUES (?, ?, ?) "
                "ON CONFLICT(chave) DO UPDATE SET valor = excluded.valor, "
                "atualizado_em = excluded.atualizado_em",
                (chave, _dump(valor), datetime.datetime.now().isoformat()),
            )
        _confirmar_salvo(f"config:{chave}", valor)
        _marcar_sucesso()
    except Exception as erro:
        _registrar_erro(f"Falha ao salvar a configuração '{chave}'", erro)


def salvar_bancas(bancas_config: dict[str, list[str]], forcar: bool = False) -> None:
    if not forcar and not _mudou("bancas", bancas_config):
        return
    anterior: dict[str, list[str]] = st.session_state.get("_salvo_bancas", {})
    try:
        conexao = conectar()
        with _LOCK, conexao:
            for ordem_banca, (banca, locais) in enumerate(bancas_config.items()):
                conexao.execute(
                    "INSERT INTO bancas (nome, ordem) VALUES (?, ?)"
                    " ON CONFLICT(nome) DO UPDATE SET ordem = excluded.ordem",
                    (banca, ordem_banca),
                )
                conexao.executemany(
                    "INSERT INTO localidades (banca, nome, ordem) VALUES (?, ?, ?)"
                    " ON CONFLICT(banca, nome) DO UPDATE SET ordem = excluded.ordem",
                    [(banca, local, i) for i, local in enumerate(locais)],
                )
                for removido in set(anterior.get(banca, [])) - set(locais):
                    conexao.execute(
                        "DELETE FROM localidades WHERE banca = ? AND nome = ?",
                        (banca, removido),
                    )
            for banca_removida in set(anterior) - set(bancas_config):
                conexao.execute("DELETE FROM bancas WHERE nome = ?", (banca_removida,))
                conexao.execute(
                    "DELETE FROM localidades WHERE banca = ?", (banca_removida,)
                )
        st.session_state["_salvo_bancas"] = {
            b: list(l) for b, l in bancas_config.items()
        }
        _confirmar_salvo("bancas", bancas_config)
        _marcar_sucesso()
    except Exception as erro:
        _registrar_erro("Falha ao salvar bancas e localidades", erro)


def _sem_id(viagem: Viagem) -> dict[str, Any]:
    return {k: v for k, v in viagem.items() if k != "id"}


def _parametros_viagem(viagem: Viagem) -> tuple:
    inicio = para_data(viagem.get("Data Inicio"))
    fim = para_data(viagem.get("Data Fim"))
    if not inicio or not fim:
        raise ValueError(
            f"Viagem para {viagem.get('Destino', '?')} sem data de início ou fim."
        )
    return (
        viagem.get("Banca", ""),
        _inteiro(viagem.get("Numero Banca Itinerante"), 1),
        viagem.get("Destino", ""),
        inicio.isoformat(),
        fim.isoformat(),
        viagem.get("Turno", TURNO_INTEGRAL),
        _dump(viagem.get("Turnos Por Data") or {}),
        _inteiro(viagem.get("Examinadores")),
        _dump(viagem.get("Bancas Apoio") or []),
        _dump(viagem.get("Examinadores Por Banca") or {}),
        str(viagem.get("Observações", "")),
        str(viagem.get("Horario Inicio") or ""),
        str(viagem.get("Horario Fim") or ""),
    )


_COLUNAS_VIAGEM = (
    "banca, numero_equipe, destino, data_inicio, data_fim, turno,"
    " turnos_por_data, examinadores, bancas_apoio, examinadores_por_banca,"
    " observacoes, horario_inicio, horario_fim"
)
_MARCADORES_VIAGEM = ", ".join(["?"] * 13)


def salvar_viagens(viagens: list[Viagem], forcar: bool = False) -> None:
    """Grava só as viagens novas, alteradas ou removidas por esta sessão.

    Viagens cadastradas por outra pessoa depois que esta sessão carregou os
    dados não são tocadas (antes eram apagadas pelo DELETE geral).
    """
    comparavel = [_sem_id(v) for v in viagens]
    if not forcar and not _mudou("viagens", comparavel):
        return
    conhecidas: dict[int, str] = st.session_state.get("_salvo_viagens", {})
    try:
        conexao = conectar()
        novas_assinaturas: dict[int, str] = {}
        with _LOCK, conexao:
            presentes: set[int] = set()
            for viagem in viagens:
                parametros = _parametros_viagem(viagem)
                assinatura = _assinatura(_sem_id(viagem))
                id_viagem = viagem.get("id")
                if id_viagem is None:
                    cursor = conexao.execute(
                        f"INSERT INTO viagens ({_COLUNAS_VIAGEM})"
                        f" VALUES ({_MARCADORES_VIAGEM})",
                        parametros,
                    )
                    viagem["id"] = cursor.lastrowid
                elif forcar or conhecidas.get(id_viagem) != assinatura:
                    cursor = conexao.execute(
                        "UPDATE viagens SET banca = ?, numero_equipe = ?, destino = ?,"
                        " data_inicio = ?, data_fim = ?, turno = ?, turnos_por_data = ?,"
                        " examinadores = ?, bancas_apoio = ?, examinadores_por_banca = ?,"
                        " observacoes = ?, horario_inicio = ?, horario_fim = ?"
                        " WHERE id = ?",
                        parametros + (id_viagem,),
                    )
                    if cursor.rowcount == 0:
                        conexao.execute(
                            f"INSERT INTO viagens (id, {_COLUNAS_VIAGEM})"
                            f" VALUES (?, {_MARCADORES_VIAGEM})",
                            (id_viagem,) + parametros,
                        )
                presentes.add(viagem["id"])
                novas_assinaturas[viagem["id"]] = assinatura
            for removida in set(conhecidas) - presentes:
                conexao.execute("DELETE FROM viagens WHERE id = ?", (removida,))
        st.session_state["_salvo_viagens"] = novas_assinaturas
        _confirmar_salvo("viagens", comparavel)
        _marcar_sucesso()
    except Exception as erro:
        _registrar_erro("Falha ao salvar as viagens", erro)


_SQL_UPSERT_HISTORICO = (
    "INSERT INTO historico (chave, data, horario, dia_semana, status,"
    " exam_m, exam_t, vagas, total, bloqueio_auto)"
    " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    " ON CONFLICT(chave, data, horario) DO UPDATE SET"
    " dia_semana = excluded.dia_semana,"
    " status = excluded.status,"
    " exam_m = excluded.exam_m,"
    " exam_t = excluded.exam_t,"
    " vagas = excluded.vagas,"
    " total = excluded.total,"
    " bloqueio_auto = excluded.bloqueio_auto"
)


def salvar_historico(chave: str, df: pd.DataFrame, forcar: bool = False) -> None:
    """Grava as linhas informadas por upsert, sem apagar outras linhas da
    mesma chave que não estejam neste DataFrame.

    O calendário em tela é reconstruído a partir dos parâmetros atuais (dias
    da semana, horários, feriados). Se um parâmetro muda, a reconstrução pode
    gerar menos linhas; um DELETE + INSERT apagaria os lançamentos que
    ficaram de fora. Com upsert, o que não está aqui não é tocado e volta a
    aparecer se o parâmetro voltar ao que era. Essas linhas "fora da grade"
    também não entram em nenhum total (ver ``filtrar_grade_vigente``).
    """
    if df is None or df.empty:
        return
    registros = df.to_dict(orient="records")
    if not forcar and not _mudou(f"historico:{chave}", registros):
        return
    try:
        conexao = conectar()
        linhas = _linhas_historico_para_gravar(chave, registros)
        with _LOCK, conexao:
            conexao.executemany(_SQL_UPSERT_HISTORICO, linhas)
        _confirmar_salvo(f"historico:{chave}", registros)
        _marcar_sucesso()
    except Exception as erro:
        _registrar_erro(f"Falha ao salvar o calendário '{chave}'", erro)


def _linhas_historico_para_gravar(
    chave: str, registros: list[dict[str, Any]]
) -> list[tuple]:
    linhas = []
    for registro in registros:
        vagas = {c: _inteiro(registro.get(c, 0)) for c in COLS_VAGAS}
        linhas.append(
            (
                chave,
                str(registro.get("Data", "")),
                str(registro.get("Horário", "")),
                str(registro.get("Dia da Semana", "")),
                str(registro.get("Status", STATUS_DISPONIVEL)),
                _inteiro(registro.get("Exam. M")),
                _inteiro(registro.get("Exam. T")),
                _dump(vagas),
                _inteiro(registro.get("Total")),
                1 if _inteiro(registro.get(COL_AUTO)) else 0,
            )
        )
    return linhas


def _salvar_mapa(nome: str, dados: dict[str, Any], forcar: bool = False) -> None:
    """Upsert das chaves alteradas e DELETE só das chaves que esta sessão
    removeu."""
    if not forcar and not _mudou(f"mapa:{nome}", dados):
        return
    tabela, coluna, eh_json = MAPAS[nome]
    salvos = st.session_state.setdefault("_salvo_mapas", {})
    conhecidos: dict[str, str] = salvos.get(nome, {})
    atuais = {chave: _assinatura(valor) for chave, valor in dados.items()}
    try:
        conexao = conectar()
        with _LOCK, conexao:
            for chave, valor in dados.items():
                if forcar or conhecidos.get(chave) != atuais[chave]:
                    conexao.execute(
                        f"INSERT INTO {tabela} (chave, {coluna}) VALUES (?, ?)"
                        f" ON CONFLICT(chave) DO UPDATE SET {coluna} = excluded.{coluna}",
                        (chave, _dump(valor) if eh_json else valor),
                    )
            for chave in set(conhecidos) - set(dados):
                conexao.execute(f"DELETE FROM {tabela} WHERE chave = ?", (chave,))
        salvos[nome] = atuais
        _confirmar_salvo(f"mapa:{nome}", dados)
        _marcar_sucesso()
    except Exception as erro:
        _registrar_erro(f"Falha ao salvar a tabela '{tabela}'", erro)


def salvar_dias_permitidos(dados: dict[str, list[str]], forcar: bool = False) -> None:
    _salvar_mapa("dias_permitidos", dados, forcar)


def salvar_feriados_locais(dados: dict[str, str], forcar: bool = False) -> None:
    _salvar_mapa("feriados_locais", dados, forcar)


def salvar_disponibilidade(dados: dict[str, int], forcar: bool = False) -> None:
    _salvar_mapa("disponibilidade", dados, forcar)


def salvar_parametros_local(dados: dict[str, dict], forcar: bool = False) -> None:
    _salvar_mapa("parametros_local", dados, forcar)


def salvar_demandas(dados: dict[str, dict], forcar: bool = False) -> None:
    _salvar_mapa("demandas", dados, forcar)


def salvar_efetivo_dia(dados: dict[str, int], forcar: bool = False) -> None:
    _salvar_mapa("efetivo_dia", dados, forcar)


def remover_historico_do_local(banca: str, local: str) -> None:
    """Limpeza em cascata do histórico e das demandas de uma localidade.

    As chaves são buscadas no banco (e não só na sessão), para pegar também
    calendários criados por outra pessoa depois que esta sessão carregou.
    """
    def pertence(chave: str) -> bool:
        partes = desmontar_chave_historico(chave)
        return bool(partes) and partes[0] == banca and partes[3] == local

    try:
        conexao = conectar()
        with _LOCK, conexao:
            alvos = {
                linha["chave"]
                for linha in conexao.execute("SELECT DISTINCT chave FROM historico")
                if pertence(linha["chave"])
            }
            alvos |= {c for c in st.session_state.get("historico_localidades", {}) if pertence(c)}
            for chave in alvos:
                conexao.execute("DELETE FROM historico WHERE chave = ?", (chave,))
        for chave in alvos:
            st.session_state["historico_localidades"].pop(chave, None)
            st.session_state.get("_assinaturas", {}).pop(f"historico:{chave}", None)
        demandas = st.session_state.get("demandas_dict", {})
        for chave in [c for c in demandas if pertence(c)]:
            demandas.pop(chave, None)
        salvar_demandas(demandas)
        _marcar_sucesso()
    except Exception as erro:
        _registrar_erro(f"Falha ao remover o histórico de '{local}'", erro)


# --- Carga inicial ----------------------------------------------------------
CONFIG_PADRAO: dict[str, Any] = {
    "lista_horarios": HORARIOS_PADRAO,
    "horarios_inativos": [],
    "capacidade_bancas": CAPACIDADE_BANCAS_PADRAO,
    "limite_fora_sede_bancas": LIMITE_FORA_SEDE_PADRAO,
    "controle_capacidade_ativo": True,
    "horarios_fim_turno": HORARIOS_FIM_TURNO_PADRAO,
    "facultativos_bloqueados": FACULTATIVOS_PADRAO_BLOQUEADOS,
}


def _semear_padroes(conexao: sqlite3.Connection) -> None:
    """Semeia a estrutura padrão só em banco novo.

    Antes, o critério era "não há bancas": quem apagasse todas as bancas via
    as seis padrão voltarem no próximo acesso, com horários e capacidades
    sobrescritos (INSERT OR REPLACE). Agora o marco é a existência de
    ``schema_version`` e nada que já exista é sobrescrito.
    """
    with _LOCK:
        ja_semeado = conexao.execute(
            "SELECT valor FROM config WHERE chave = 'schema_version'"
        ).fetchone()
        tem_bancas = conexao.execute("SELECT COUNT(*) AS n FROM bancas").fetchone()["n"]
        agora = datetime.datetime.now().isoformat()
        try:
            versao_anterior = int(json.loads(ja_semeado["valor"])) if ja_semeado else None
        except (TypeError, ValueError, json.JSONDecodeError):
            versao_anterior = 0
        with conexao:
            if versao_anterior is not None and versao_anterior < 8 and tem_bancas:
                _alinhar_cadastro_oficial(conexao)
            if not ja_semeado and not tem_bancas:
                LOG.info("Banco novo: semeando estrutura padrão")
                for ordem_banca, (banca, locais) in enumerate(BANCAS_PADRAO.items()):
                    conexao.execute(
                        "INSERT OR IGNORE INTO bancas (nome, ordem) VALUES (?, ?)",
                        (banca, ordem_banca),
                    )
                    conexao.executemany(
                        "INSERT OR IGNORE INTO localidades (banca, nome, ordem)"
                        " VALUES (?, ?, ?)",
                        [(banca, local, i) for i, local in enumerate(locais)],
                    )
                for chave, pistas in PISTAS_MOTO_PADRAO.items():
                    conexao.execute(
                        "INSERT OR IGNORE INTO parametros_local (chave, valor)"
                        " VALUES (?, ?)",
                        (chave, _dump({"pistas_moto": pistas})),
                    )
                for chave, valor in CONFIG_PADRAO.items():
                    conexao.execute(
                        "INSERT OR IGNORE INTO config (chave, valor, atualizado_em)"
                        " VALUES (?, ?, ?)",
                        (chave, _dump(valor), agora),
                    )
            conexao.execute(
                "INSERT INTO config (chave, valor, atualizado_em) VALUES"
                " ('schema_version', ?, ?) ON CONFLICT(chave) DO UPDATE SET"
                " valor = excluded.valor, atualizado_em = excluded.atualizado_em",
                (_dump(SCHEMA_VERSION), agora),
            )


def _alinhar_cadastro_oficial(conexao: sqlite3.Connection) -> None:
    """Migração 8: leva bancos antigos ao cadastro oficial de municípios.

    Roda dentro da transação de ``_semear_padroes``. Só mexe em bancas que
    existem (banca apagada de propósito não volta) e nunca apaga dado
    lançado:
      * renomeia os postos que mudaram de nome (ex.: "São Luís Pátio" →
        "São Luís - Pátio DETRAN") em todas as tabelas que guardam o nome;
      * acrescenta os municípios que faltam;
      * remove as localidades que a versão 4 semeava e que não fazem parte do
        cadastro, desde que nunca tenham recebido vagas nem viagens;
      * reordena as localidades (região metropolitana/sede primeiro).
    """
    bancas = {linha["nome"] for linha in conexao.execute("SELECT nome FROM bancas")}

    def locais_da(banca: str) -> set[str]:
        return {
            linha["nome"]
            for linha in conexao.execute(
                "SELECT nome FROM localidades WHERE banca = ?", (banca,)
            )
        }

    def chaves_historico(banca: str, local: str) -> list[str]:
        alvo = []
        for linha in conexao.execute("SELECT DISTINCT chave FROM historico"):
            partes = desmontar_chave_historico(linha["chave"])
            if partes and partes[0] == banca and partes[3] == local:
                alvo.append(linha["chave"])
        return alvo

    for banca, renomeacoes in RENOMEACOES_V5.items():
        if banca not in bancas:
            continue
        for antigo, novo in renomeacoes.items():
            existentes = locais_da(banca)
            if antigo not in existentes or novo in existentes:
                continue
            LOG.info("Migração 8: %s / %s → %s", banca, antigo, novo)
            conexao.execute(
                "UPDATE localidades SET nome = ? WHERE banca = ? AND nome = ?",
                (novo, banca, antigo),
            )
            for chave in chaves_historico(banca, antigo):
                partes = desmontar_chave_historico(chave)
                nova = chave_historico(partes[0], partes[1], partes[2], novo)
                conexao.execute("UPDATE historico SET chave = ? WHERE chave = ?", (nova, chave))
                conexao.execute(
                    "UPDATE OR IGNORE demandas SET chave = ? WHERE chave = ?", (nova, chave)
                )
            for tabela in ("parametros_local", "dias_permitidos", "feriados_locais"):
                conexao.execute(
                    f"UPDATE OR IGNORE {tabela} SET chave = ? WHERE chave = ?",
                    (chave_local(banca, novo), chave_local(banca, antigo)),
                )
            conexao.execute(
                "UPDATE viagens SET destino = ? WHERE banca = ? AND destino = ?",
                (novo, banca, antigo),
            )

    for banca, obsoletas in OBSOLETAS_V5.items():
        if banca not in bancas:
            continue
        for local in obsoletas:
            if local not in locais_da(banca):
                continue
            chaves = chaves_historico(banca, local)
            com_vagas = any(
                conexao.execute(
                    "SELECT 1 FROM historico WHERE chave = ? AND total > 0 LIMIT 1", (chave,)
                ).fetchone()
                for chave in chaves
            )
            com_viagem = conexao.execute(
                "SELECT 1 FROM viagens WHERE banca = ? AND destino = ? LIMIT 1", (banca, local)
            ).fetchone()
            if com_vagas or com_viagem:
                LOG.info("Migração 8: %s / %s mantida (tem dados)", banca, local)
                continue
            LOG.info("Migração 8: %s / %s removida (fora do cadastro, sem dados)", banca, local)
            conexao.execute(
                "DELETE FROM localidades WHERE banca = ? AND nome = ?", (banca, local)
            )
            for chave in chaves:
                conexao.execute("DELETE FROM historico WHERE chave = ?", (chave,))
                conexao.execute("DELETE FROM demandas WHERE chave = ?", (chave,))
            for tabela in ("parametros_local", "dias_permitidos", "feriados_locais"):
                conexao.execute(
                    f"DELETE FROM {tabela} WHERE chave = ?", (chave_local(banca, local),)
                )

    for banca, oficiais in BANCAS_PADRAO.items():
        if banca not in bancas:
            continue
        conexao.executemany(
            "INSERT OR IGNORE INTO localidades (banca, nome, ordem) VALUES (?, ?, 0)",
            [(banca, local) for local in oficiais],
        )
        ordenadas = ordenar_localidades_da_banca(banca, sorted(locais_da(banca)))
        conexao.executemany(
            "UPDATE localidades SET ordem = ? WHERE banca = ? AND nome = ?",
            [(i, banca, local) for i, local in enumerate(ordenadas)],
        )


def _ler_tudo(conexao: sqlite3.Connection) -> dict[str, Any]:
    """Lê o banco inteiro dentro do lock (a conexão é compartilhada entre as
    sessões, e leitura simultânea a uma gravação na mesma conexão falha)."""
    with _LOCK:
        return {
            "bancas_config": _ler_bancas(conexao),
            "config": {
                chave: _ler_config(conexao, chave, padrao)
                for chave, padrao in CONFIG_PADRAO.items()
            },
            "viagens": _ler_viagens(conexao),
            "historico": _ler_historico(conexao),
            "mapas": {nome: _ler_mapa(conexao, nome) for nome in MAPAS},
        }


def carregar_estado(forcar: bool = False) -> None:
    """Carrega o banco para o session_state uma vez por sessão.

    Uma sessão aberta antes de uma atualização do sistema (outra versão do
    esquema) recarrega tudo: senão ficaria sem as chaves novas do estado.
    """
    if (
        st.session_state.get("_estado_carregado")
        and st.session_state.get("_versao_estado") == SCHEMA_VERSION
        and not forcar
    ):
        return

    conexao = conectar()
    _semear_padroes(conexao)
    dados = _ler_tudo(conexao)
    config = dados["config"]
    mapas = dados["mapas"]

    st.session_state["bancas_config"] = dados["bancas_config"]
    st.session_state["lista_horarios"] = list(config["lista_horarios"] or HORARIOS_PADRAO)
    st.session_state["horarios_inativos"] = list(config["horarios_inativos"] or [])
    st.session_state["capacidade_bancas"] = {
        **CAPACIDADE_BANCAS_PADRAO,
        **(config["capacidade_bancas"] or {}),
    }
    st.session_state["limite_fora_sede_bancas"] = {
        **LIMITE_FORA_SEDE_PADRAO,
        **(config["limite_fora_sede_bancas"] or {}),
    }
    st.session_state["controle_capacidade_ativo"] = bool(config["controle_capacidade_ativo"])
    st.session_state["horarios_fim_turno"] = list(config["horarios_fim_turno"] or [])
    st.session_state["facultativos_bloqueados"] = list(config["facultativos_bloqueados"] or [])
    st.session_state["viagens_registradas"] = garantir_numeracao_equipes(dados["viagens"])
    st.session_state["historico_localidades"] = dados["historico"]
    st.session_state["dias_permitidos_dict"] = mapas["dias_permitidos"]
    st.session_state["feriados_locais_dict"] = mapas["feriados_locais"]
    st.session_state["disponibilidade_semanal"] = {
        chave: _inteiro(valor) for chave, valor in mapas["disponibilidade"].items()
    }
    st.session_state["parametros_local_dict"] = mapas["parametros_local"]
    st.session_state["demandas_dict"] = mapas["demandas"]
    st.session_state["efetivo_dia_dict"] = {
        chave: _inteiro(valor) for chave, valor in mapas["efetivo_dia"].items()
    }

    # Registra o que acabou de ser lido como "já salvo" (base das gravações
    # diferenciais e do teste de mudança).
    st.session_state["_assinaturas"] = {}
    st.session_state["_salvo_bancas"] = {
        b: list(l) for b, l in st.session_state["bancas_config"].items()
    }
    _confirmar_salvo("bancas", st.session_state["bancas_config"])
    st.session_state["_salvo_viagens"] = {
        v["id"]: _assinatura(_sem_id(v)) for v in st.session_state["viagens_registradas"]
    }
    _confirmar_salvo(
        "viagens", [_sem_id(v) for v in st.session_state["viagens_registradas"]]
    )
    st.session_state["_salvo_mapas"] = {}
    for nome, chave_estado in (
        ("dias_permitidos", "dias_permitidos_dict"),
        ("feriados_locais", "feriados_locais_dict"),
        ("disponibilidade", "disponibilidade_semanal"),
        ("parametros_local", "parametros_local_dict"),
        ("demandas", "demandas_dict"),
        ("efetivo_dia", "efetivo_dia_dict"),
    ):
        valor = st.session_state[chave_estado]
        st.session_state["_salvo_mapas"][nome] = {
            k: _assinatura(v) for k, v in valor.items()
        }
        _confirmar_salvo(f"mapa:{nome}", valor)
    for chave in CONFIG_PADRAO:
        _confirmar_salvo(f"config:{chave}", config[chave])
    for chave, df in st.session_state["historico_localidades"].items():
        _confirmar_salvo(f"historico:{chave}", df.to_dict(orient="records"))

    # Muda o sufixo das chaves dos widgets ligados a dados: depois de uma
    # recarga ou restauração, nenhum widget fica com o valor antigo preso.
    st.session_state["_versao_widgets"] = st.session_state.get("_versao_widgets", 0) + 1
    st.session_state["_estado_carregado"] = True
    st.session_state["_versao_estado"] = SCHEMA_VERSION
    st.session_state.setdefault("erro_persistencia", None)
    LOG.info(
        "Estado carregado: %d bancas, %d viagens, %d calendários",
        len(st.session_state["bancas_config"]),
        len(st.session_state["viagens_registradas"]),
        len(st.session_state["historico_localidades"]),
    )


def recarregar_do_banco() -> None:
    carregar_estado(forcar=True)


def sufixo_widget() -> str:
    """Sufixo para chaves de widgets que exibem dados persistidos."""
    return f"v{st.session_state.get('_versao_widgets', 0)}"


# --- Backup -----------------------------------------------------------------
def montar_backup() -> dict[str, Any]:
    viagens = []
    for viagem in st.session_state.get("viagens_registradas", []):
        copia = dict(viagem)
        copia.pop("id", None)
        copia["Data Inicio"] = para_data(viagem.get("Data Inicio")).isoformat()
        copia["Data Fim"] = para_data(viagem.get("Data Fim")).isoformat()
        viagens.append(copia)

    historico = {
        chave: df.fillna(0).to_dict(orient="records")
        for chave, df in st.session_state.get("historico_localidades", {}).items()
    }

    return {
        "schema_version": SCHEMA_VERSION,
        "gerado_em": datetime.datetime.now().isoformat(),
        "bancas_config": st.session_state.get("bancas_config", {}),
        "lista_horarios": st.session_state.get("lista_horarios", []),
        "horarios_inativos": st.session_state.get("horarios_inativos", []),
        "historico_localidades": historico,
        "feriados_locais_dict": st.session_state.get("feriados_locais_dict", {}),
        "viagens_registradas": viagens,
        "dias_permitidos_dict": st.session_state.get("dias_permitidos_dict", {}),
        "capacidade_bancas": st.session_state.get("capacidade_bancas", {}),
        "limite_fora_sede_bancas": st.session_state.get("limite_fora_sede_bancas", {}),
        "disponibilidade_semanal": st.session_state.get("disponibilidade_semanal", {}),
        "controle_capacidade_ativo": st.session_state.get(
            "controle_capacidade_ativo", False
        ),
        "horarios_fim_turno": st.session_state.get("horarios_fim_turno", []),
        "facultativos_bloqueados": st.session_state.get("facultativos_bloqueados", []),
        "parametros_local_dict": st.session_state.get("parametros_local_dict", {}),
        "demandas_dict": st.session_state.get("demandas_dict", {}),
        "efetivo_dia_dict": st.session_state.get("efetivo_dia_dict", {}),
    }


def backup_em_json() -> str:
    return _dump(montar_backup())


CAMPOS_OBRIGATORIOS = {
    "bancas_config": dict,
    "viagens_registradas": list,
    "historico_localidades": dict,
}


def validar_backup(dados: Any) -> dict[str, Any]:
    """Valida a estrutura antes de aplicar. Levanta ``BackupInvalido``."""
    if not isinstance(dados, dict):
        raise BackupInvalido("O arquivo não contém um objeto JSON no nível raiz.")

    versao = dados.get("schema_version")
    if versao is not None and not isinstance(versao, int):
        raise BackupInvalido("O campo 'schema_version' deveria ser um número inteiro.")
    if isinstance(versao, int) and versao > SCHEMA_VERSION:
        raise BackupInvalido(
            f"Backup gerado por uma versão mais nova do sistema (schema {versao};"
            f" esta versão entende até {SCHEMA_VERSION})."
        )

    for campo, tipo in CAMPOS_OBRIGATORIOS.items():
        if campo not in dados:
            raise BackupInvalido(f"Campo obrigatório ausente: '{campo}'.")
        if not isinstance(dados[campo], tipo):
            raise BackupInvalido(
                f"O campo '{campo}' deveria ser {tipo.__name__}, "
                f"veio {type(dados[campo]).__name__}."
            )

    for banca, locais in dados["bancas_config"].items():
        if not isinstance(banca, str) or not isinstance(locais, list):
            raise BackupInvalido(
                "Cada banca em 'bancas_config' precisa mapear para uma lista de"
                " localidades."
            )

    for indice, viagem in enumerate(dados["viagens_registradas"], 1):
        if not isinstance(viagem, dict):
            raise BackupInvalido(f"A viagem nº {indice} não é um objeto.")
        for campo in ("Banca", "Destino", "Data Inicio", "Data Fim"):
            if campo not in viagem:
                raise BackupInvalido(f"A viagem nº {indice} está sem o campo '{campo}'.")
        if para_data(viagem["Data Inicio"]) is None:
            raise BackupInvalido(
                f"A viagem nº {indice} tem data de início ilegível:"
                f" {viagem['Data Inicio']!r}."
            )
        if para_data(viagem["Data Fim"]) is None:
            raise BackupInvalido(
                f"A viagem nº {indice} tem data de término ilegível:"
                f" {viagem['Data Fim']!r}."
            )

    for chave, registros in dados["historico_localidades"].items():
        if not isinstance(registros, list):
            raise BackupInvalido(f"O calendário '{chave}' não é uma lista de linhas.")
        if desmontar_chave_historico(chave) is None:
            raise BackupInvalido(
                f"A chave de calendário '{chave}' está fora do formato esperado"
                " (banca||mês||ano||localidade)."
            )
    return dados


def _migrar_viagem_antiga(viagem: dict[str, Any]) -> Viagem:
    """Converte o formato 'Apoio Conjunto' booleano para a lista explícita."""
    convertida = dict(viagem)
    convertida.pop("id", None)
    convertida["Data Inicio"] = para_data(viagem.get("Data Inicio"))
    convertida["Data Fim"] = para_data(viagem.get("Data Fim"))
    convertida.setdefault("Turnos Por Data", {})

    if "Bancas Apoio" not in convertida:
        convertida["Bancas Apoio"] = []
    if "Examinadores Por Banca" not in convertida:
        por_banca: dict[str, int] = {}
        if viagem.get("Apoio Conjunto"):
            for banca, campo in (
                ("Banca Caxias", "Examinadores Caxias"),
                ("Banca Timon", "Examinadores Timon"),
            ):
                if campo in viagem:
                    por_banca[banca] = _inteiro(viagem[campo])
            convertida["Bancas Apoio"] = [
                b for b in por_banca if b != convertida.get("Banca")
            ]
        convertida["Examinadores Por Banca"] = por_banca
    for campo in ("Apoio Conjunto", "Examinadores Caxias", "Examinadores Timon", "Período"):
        convertida.pop(campo, None)
    return convertida


def _gravar_backup_no_banco(conexao: sqlite3.Connection, dados: dict[str, Any]) -> None:
    """Substitui o conteúdo do banco pelo backup, em UMA transação.

    Se qualquer passo falhar, nada muda (rollback). Antes, cada tabela era
    gravada separadamente e calendários que existiam no banco mas não no
    backup sobreviviam à restauração e reapareciam na recarga.
    """
    agora = datetime.datetime.now().isoformat()
    viagens = garantir_numeracao_equipes(
        [_migrar_viagem_antiga(v) for v in dados["viagens_registradas"]]
    )
    with conexao:
        for tabela in (
            "bancas",
            "localidades",
            "viagens",
            "historico",
            "dias_permitidos",
            "feriados_locais",
            "disponibilidade_semanal",
            "parametros_local",
            "demandas",
            "efetivo_diario",
        ):
            conexao.execute(f"DELETE FROM {tabela}")
        for ordem_banca, (banca, locais) in enumerate(dados["bancas_config"].items()):
            conexao.execute(
                "INSERT INTO bancas (nome, ordem) VALUES (?, ?)", (banca, ordem_banca)
            )
            conexao.executemany(
                "INSERT OR IGNORE INTO localidades (banca, nome, ordem) VALUES (?, ?, ?)",
                [(banca, str(local), i) for i, local in enumerate(locais)],
            )
        for viagem in viagens:
            conexao.execute(
                f"INSERT INTO viagens ({_COLUNAS_VIAGEM}) VALUES ({_MARCADORES_VIAGEM})",
                _parametros_viagem(viagem),
            )
        for chave, registros in dados["historico_localidades"].items():
            if registros:
                conexao.executemany(
                    _SQL_UPSERT_HISTORICO, _linhas_historico_para_gravar(chave, registros)
                )
        mapas_backup = {
            "dias_permitidos": dados.get("dias_permitidos_dict") or {},
            "feriados_locais": dados.get("feriados_locais_dict") or {},
            "disponibilidade": {
                k: _inteiro(v) for k, v in (dados.get("disponibilidade_semanal") or {}).items()
            },
            "parametros_local": dados.get("parametros_local_dict")
            or {k: {"pistas_moto": v} for k, v in PISTAS_MOTO_PADRAO.items()},
            "demandas": dados.get("demandas_dict") or {},
            "efetivo_dia": {
                k: _inteiro(v) for k, v in (dados.get("efetivo_dia_dict") or {}).items()
            },
        }
        for nome, valores in mapas_backup.items():
            tabela, coluna, eh_json = MAPAS[nome]
            conexao.executemany(
                f"INSERT INTO {tabela} (chave, {coluna}) VALUES (?, ?)",
                [(k, _dump(v) if eh_json else v) for k, v in valores.items()],
            )
        configuracoes = {
            "lista_horarios": dados.get("lista_horarios") or list(HORARIOS_PADRAO),
            "horarios_inativos": dados.get("horarios_inativos") or [],
            "capacidade_bancas": {
                **CAPACIDADE_BANCAS_PADRAO,
                **(dados.get("capacidade_bancas") or {}),
            },
            "limite_fora_sede_bancas": {
                **LIMITE_FORA_SEDE_PADRAO,
                **(dados.get("limite_fora_sede_bancas") or {}),
            },
            "controle_capacidade_ativo": bool(dados.get("controle_capacidade_ativo", True)),
            "horarios_fim_turno": dados.get("horarios_fim_turno", HORARIOS_FIM_TURNO_PADRAO),
            "facultativos_bloqueados": dados.get(
                "facultativos_bloqueados", FACULTATIVOS_PADRAO_BLOQUEADOS
            ),
        }
        for chave, valor in configuracoes.items():
            conexao.execute(
                "INSERT INTO config (chave, valor, atualizado_em) VALUES (?, ?, ?)"
                " ON CONFLICT(chave) DO UPDATE SET valor = excluded.valor,"
                " atualizado_em = excluded.atualizado_em",
                (chave, _dump(valor), agora),
            )
        # Backup de versão anterior: aplica a mesma migração de cadastro.
        if _inteiro(dados.get("schema_version")) < 8:
            _alinhar_cadastro_oficial(conexao)


def aplicar_backup(dados: dict[str, Any]) -> None:
    """Valida, tira um snapshot de segurança e restaura tudo ou nada."""
    validar_backup(dados)
    if salvar_snapshot(motivo="antes_de_restaurar") is None:
        raise RuntimeError(
            "Não foi possível gravar o snapshot de segurança; a restauração foi"
            " cancelada para não arriscar os dados atuais."
        )
    conexao = conectar()
    try:
        with _LOCK:
            _gravar_backup_no_banco(conexao, dados)
    except Exception as erro:
        _registrar_erro("Falha ao restaurar o backup (nada foi alterado)", erro)
        raise
    carregar_estado(forcar=True)
    _marcar_sucesso()


def salvar_snapshot(motivo: str = "manual") -> Path | None:
    """Snapshot rotativo em disco, mantendo os últimos N arquivos."""
    try:
        pasta = pasta_snapshots()
        pasta.mkdir(parents=True, exist_ok=True)
        carimbo = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = pasta / f"snapshot_{carimbo}_{motivo}.json"
        destino.write_text(backup_em_json(), encoding="utf-8")

        existentes = sorted(pasta.glob("snapshot_*.json"))
        for antigo in existentes[:-MAX_SNAPSHOTS]:
            antigo.unlink(missing_ok=True)
        LOG.info("Snapshot gravado em %s", destino)
        return destino
    except Exception as erro:
        _registrar_erro("Falha ao gravar o snapshot de segurança", erro)
        return None


# ===========================================================================
# 4. EXPORTADORES
# ===========================================================================

LARGURA_UTIL_A4_PAISAGEM = landscape(A4)[0] - 30  # margens de 15pt


def _estilos_pdf() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "titulo": ParagraphStyle(
            "TituloOficial",
            parent=base["Heading1"],
            fontSize=12,
            alignment=1,
            spaceAfter=4,
        ),
        "subtitulo": ParagraphStyle(
            "SubtituloOficial",
            parent=base["Normal"],
            fontSize=9,
            alignment=1,
            spaceAfter=10,
        ),
        "grupo": ParagraphStyle(
            "GrupoLocalidades",
            parent=base["Heading2"],
            fontSize=10,
            textColor=colors.HexColor(COR_PRIMARIA),
            spaceBefore=10,
            spaceAfter=4,
        ),
        "local": ParagraphStyle(
            "NomeLocalidade",
            parent=base["Heading3"],
            fontSize=9,
            textColor=colors.HexColor(COR_SECUNDARIA),
            spaceBefore=6,
            spaceAfter=3,
        ),
        "nota": ParagraphStyle(
            "NotaRodape", parent=base["Normal"], fontSize=7.5, textColor=colors.grey
        ),
        "corpo": ParagraphStyle("Corpo", parent=base["Normal"], fontSize=8),
        "corpo_branco": ParagraphStyle(
            "CorpoBranco", parent=base["Normal"], fontSize=7.5, textColor=colors.white
        ),
    }


def _cabecalho_oficial(elementos: list, estilos: dict, subtitulo: str) -> None:
    elementos.append(
        Paragraph(
            "<b>DETRAN-MA — DEPARTAMENTO ESTADUAL DE TRÂNSITO DO MARANHÃO</b>",
            estilos["titulo"],
        )
    )
    elementos.append(Paragraph(subtitulo, estilos["subtitulo"]))


def _estilo_tabela_padrao(linhas_destaque: Sequence[int] = ()) -> TableStyle:
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COR_PRIMARIA)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
    ]
    for indice in linhas_destaque:
        estilo.extend(
            [
                ("BACKGROUND", (0, indice), (-1, indice), colors.HexColor(COR_CLARA)),
                ("FONTNAME", (0, indice), (-1, indice), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, indice), (-1, indice), colors.HexColor(COR_PRIMARIA)),
            ]
        )
    return TableStyle(estilo)


def _colunas_com_valor(df: pd.DataFrame, colunas: list[str]) -> list[str]:
    if df is None or df.empty:
        return []
    return [c for c in colunas if c in df.columns and df[c].fillna(0).sum() > 0]


def _num(valor: Any) -> int:
    return int(pd.to_numeric(valor, errors="coerce") or 0)


def _texto_ou_traco(valor: int) -> str:
    return "-" if valor == 0 else str(valor)


def _apenas_disponiveis(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    return df[(df["Status"] == STATUS_DISPONIVEL) & (df["Total"] > 0)].copy()


def _ordenar_por_data(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    copia = df.copy()
    copia["_ordem"] = pd.to_datetime(copia["Data"], format="%d/%m/%Y", errors="coerce")
    return copia.sort_values(["_ordem", "Horário"]).drop(columns=["_ordem"])


# --- PDF: grade detalhada (linha a linha por horário) -----------------------
COLUNAS_GRADE_DETALHADA = ["Data", "Dia da Semana", "Status", "Exam. M", "Exam. T", "Horário"]
ROTULOS_GRADE_DETALHADA = {"Dia da Semana": "Dia"}


def _celula_vaga(registro: dict[str, Any], coluna: str) -> str:
    """Valor da célula de vagas; C, D e E mostram o horário quebrado quando
    dividem o horário com a Cat B (ex.: '4 (08:32)')."""
    valor = _num(registro.get(coluna, 0))
    texto = _texto_ou_traco(valor)
    categoria = categoria_da_coluna(coluna)
    if valor and categoria:
        quebrado = horarios_quebrados(registro).get(categoria)
        if quebrado:
            texto += f" ({quebrado})"
    return texto


def _linhas_grade_detalhada(
    df_filtrado: pd.DataFrame, pcd_usadas: list[str]
) -> tuple[list[str], list[list[str]], list[int]]:
    """Linha a linha por horário, com subtotal por data.

    Compartilhada pelo PDF de uma única localidade e pelo PDF completo da
    banca, para que os dois documentos tragam exatamente o mesmo nível de
    detalhe (é o formato que os colaboradores usam para lançar no sistema).
    Devolve (colunas, linhas, índices das linhas de subtotal a destacar),
    com índices já contando a linha de cabeçalho.
    """
    numericas = COLS_CATEGORIA + pcd_usadas + ["Total"]
    colunas = COLUNAS_GRADE_DETALHADA + COLS_CATEGORIA + pcd_usadas + ["Total"]

    linhas: list[list[str]] = []
    destaques: list[int] = []
    for data_valor, grupo in df_filtrado.groupby("Data", sort=False):
        for registro in grupo.to_dict(orient="records"):
            linha = []
            for coluna in colunas:
                if coluna in numericas:
                    linha.append(_celula_vaga(registro, coluna))
                else:
                    linha.append(html.escape(str(registro.get(coluna, ""))))
            linhas.append(linha)

        subtotal = [f"SUBTOTAL {data_valor}", "", "", "-", "-", "-"]
        for coluna in COLS_CATEGORIA + pcd_usadas:
            subtotal.append(_texto_ou_traco(int(grupo[coluna].fillna(0).sum())))
        subtotal.append(str(int(grupo["Total"].fillna(0).sum())))
        linhas.append(subtotal)
        destaques.append(len(linhas))

    return colunas, linhas, destaques


def _rotulos(colunas: list[str]) -> list[str]:
    return [ROTULOS_GRADE_DETALHADA.get(c, c) for c in colunas]


def _larguras_grade(quantidade_colunas: int) -> list[float]:
    """Larguras da grade detalhada (A4 paisagem).

    Data, Dia, Status, Exam. M, Exam. T e Horário têm largura fixa; as
    colunas de vagas dividem o resto (mínimo 34pt, para caber '4 (08:32)').
    """
    fixas = [54, 46, 58, 34, 34, 38]
    if quantidade_colunas <= len(fixas):
        return fixas[:quantidade_colunas]
    restante = LARGURA_UTIL_A4_PAISAGEM - sum(fixas)
    demais = max(restante / (quantidade_colunas - len(fixas)), 34)
    return fixas + [demais] * (quantidade_colunas - len(fixas))


def _estilo_grade(destaques: Sequence[int]) -> TableStyle:
    """Estilo padrão + rótulo do subtotal ocupando Data, Dia e Status."""
    estilo = _estilo_tabela_padrao(destaques)
    for indice in destaques:
        estilo.add("SPAN", (0, indice), (2, indice))
        estilo.add("ALIGN", (0, indice), (2, indice), "LEFT")
    return estilo


def gerar_pdf_localidade(
    df_dados: pd.DataFrame, banca: str, local: str, mes: str, ano: int | str
) -> bytes:
    """Grade de lançamento de vagas, linha a linha por horário."""
    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15,
        title=f"Grade de vagas - {local} - {mes}/{ano}",
    )
    estilos = _estilos_pdf()
    elementos: list = []
    _cabecalho_oficial(
        elementos,
        estilos,
        "GRADE DE LANÇAMENTO DE VAGAS DE EXAMES PRÁTICOS - "
        f"{html.escape(str(banca)).upper()} ({html.escape(str(local)).upper()})"
        f" - {html.escape(str(mes))}/{html.escape(str(ano))}",
    )

    df_filtrado = _ordenar_por_data(_apenas_disponiveis(df_dados))
    if df_filtrado.empty:
        elementos.append(
            Paragraph(
                "<b>Nenhuma vaga ofertada para os parâmetros selecionados.</b>",
                estilos["subtitulo"],
            )
        )
        documento.build(elementos)
        return buffer.getvalue()

    pcd_usadas = _colunas_com_valor(df_filtrado, COLS_PCD)
    colunas, linhas, destaques = _linhas_grade_detalhada(df_filtrado, pcd_usadas)

    tabela = Table(
        [_rotulos(colunas)] + linhas, repeatRows=1, colWidths=_larguras_grade(len(colunas))
    )
    tabela.setStyle(_estilo_grade(destaques))
    elementos.append(tabela)
    elementos.append(Spacer(1, 6))
    elementos.append(Paragraph(LEGENDA_HORARIO_QUEBRADO, estilos["nota"]))
    documento.build(elementos)
    return buffer.getvalue()


LEGENDA_HORARIO_QUEBRADO = (
    "Horário entre parênteses = horário quebrado da categoria quando divide o"
    " horário com a Cat B (C +1 min, D +2 min, E +10 min)."
)


# --- PDF: calendário consolidado da banca -----------------------------------


def gerar_pdf_banca(
    dados_por_local: dict[str, pd.DataFrame],
    viagens: Sequence[Viagem],
    bancas_config: dict[str, list[str]],
    banca: str,
    mes: str,
    ano: int,
    mes_numero: int,
    efetivo_por_dia: dict[datetime.date, int],
) -> bytes:
    """Calendário completo da banca, para conferência antes da publicação.

    As localidades saem na ordem pedida: unidades da sede, região
    metropolitana e depois os demais municípios pela data do primeiro exame.
    """
    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15,
        title=f"Calendário {banca} - {mes}/{ano}",
    )
    estilos = _estilos_pdf()
    elementos: list = []
    _cabecalho_oficial(
        elementos,
        estilos,
        "CALENDÁRIO CONSOLIDADO DE EXAMES PRÁTICOS - "
        f"{html.escape(str(banca)).upper()} - {html.escape(str(mes)).upper()}/{ano}",
    )

    ordenadas = ordenar_localidades_para_pdf(banca, bancas_config, dados_por_local)
    grupo_atual = None
    total_geral = 0
    houve_conteudo = False
    primeira_localidade = True

    for grupo, local in ordenadas:
        df_local = _ordenar_por_data(_apenas_disponiveis(dados_por_local.get(local)))
        if df_local.empty:
            continue

        # Cada localidade começa em página nova: permite entregar a grade de
        # um município isoladamente ao colaborador que lança aquele
        # calendário, sem recortar o PDF.
        if not primeira_localidade:
            elementos.append(PageBreak())
        primeira_localidade = False
        houve_conteudo = True

        if grupo != grupo_atual:
            elementos.append(Paragraph(html.escape(grupo), estilos["grupo"]))
            grupo_atual = grupo

        pcd_usadas = _colunas_com_valor(df_local, COLS_PCD)
        cabecalho, linhas, destaques = _linhas_grade_detalhada(df_local, pcd_usadas)

        total_local = int(df_local["Total"].fillna(0).sum())
        total_geral += total_local

        rodape = [Paragraph(f"<b>TOTAL GERAL - {html.escape(local).upper()}</b>", estilos["corpo_branco"]), "", "", "-", "-", "-"]
        for coluna in COLS_CATEGORIA + pcd_usadas:
            rodape.append(_texto_ou_traco(int(df_local[coluna].fillna(0).sum())))
        rodape.append(str(total_local))
        linhas.append(rodape)
        indice_total_geral = len(linhas)

        dias_unicos = sorted(
            {d.day for d in (para_data(x) for x in df_local["Data"].unique()) if d}
        )
        legenda = (
            f"<b>{html.escape(local)}</b> - {len(dias_unicos)} dia(s) de exame:"
            f" {formatar_datas_exames(dias_unicos)} - {total_local} vagas"
        )

        tabela = Table(
            [_rotulos(cabecalho)] + linhas,
            repeatRows=1,
            colWidths=_larguras_grade(len(cabecalho)),
        )
        estilo_tabela = _estilo_grade(destaques + [indice_total_geral])
        estilo_tabela.add(
            "BACKGROUND",
            (0, indice_total_geral),
            (-1, indice_total_geral),
            colors.HexColor(COR_SECUNDARIA),
        )
        estilo_tabela.add(
            "TEXTCOLOR", (0, indice_total_geral), (-1, indice_total_geral), colors.white
        )
        tabela.setStyle(estilo_tabela)
        elementos.append(Paragraph(legenda, estilos["local"]))
        elementos.append(tabela)

    if not houve_conteudo:
        elementos.append(
            Paragraph(
                "<b>Nenhuma vaga lançada nesta banca para o mês selecionado.</b>",
                estilos["subtitulo"],
            )
        )
        documento.build(elementos)
        return buffer.getvalue()

    elementos.append(PageBreak())
    elementos.append(Paragraph("CONFERÊNCIA DE EFETIVO DIÁRIO", estilos["grupo"]))
    tabela_efetivo = _tabela_efetivo_mensal(
        efetivo_diario(dados_por_local, viagens, banca, ano, mes_numero), efetivo_por_dia
    )
    if tabela_efetivo:
        elementos.append(tabela_efetivo)
    elementos.append(Spacer(1, 6))
    elementos.append(
        Paragraph(
            f"Total geral de vagas da banca no mês: <b>{total_geral}</b>. "
            + LEGENDA_HORARIO_QUEBRADO
            + " Documento gerado para conferência interna em "
            f"{datetime.datetime.now().strftime('%d/%m/%Y às %H:%M')}.",
            estilos["nota"],
        )
    )

    documento.build(elementos)
    return buffer.getvalue()


def _tabela_efetivo_mensal(
    totais: dict[datetime.date, dict[str, Any]], efetivo_por_dia: dict[datetime.date, int]
) -> Table | None:
    """Manhã, tarde e pico de examinadores por dia (mesma conta da tela)."""
    if not totais:
        return None

    linhas = [["Data", "Dia", "Manhã", "Tarde", "Pico", "Efetivo", "Folga"]]
    destaques = []
    for data in sorted(totais):
        valores = totais[data]
        pico = pico_diario(valores["M"], valores["T"])
        efetivo = int(efetivo_por_dia.get(data, 0))
        linhas.append(
            [
                formatar_br(data),
                nome_dia_semana(data)[:3],
                str(valores["M"]),
                str(valores["T"]),
                str(pico),
                str(efetivo),
                str(efetivo - pico),
            ]
        )
        if pico > efetivo:
            destaques.append(len(linhas) - 1)

    tabela = Table(linhas, repeatRows=1)
    estilo = _estilo_tabela_padrao()
    for indice in destaques:
        estilo.add("BACKGROUND", (0, indice), (-1, indice), colors.HexColor("#FED7D7"))
        estilo.add("TEXTCOLOR", (0, indice), (-1, indice), colors.HexColor(COR_ALERTA))
    tabela.setStyle(estilo)
    return tabela


# --- PDF: resumo semanal do quadro de examinadores --------------------------
def gerar_pdf_semana(
    cabecalho_colunas: list[str],
    linhas_matriz: list[list[str]],
    equipes: list[dict[str, Any]],
    banca: str,
    mes: str,
    ano: int,
    rotulo_semana: str,
) -> bytes:
    """Escala semanal do quadro de examinadores, com as equipes em viagem."""
    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15,
        title=f"Escala semanal — {banca} — {rotulo_semana}",
    )
    estilos = _estilos_pdf()
    elementos: list = []
    _cabecalho_oficial(
        elementos,
        estilos,
        f"ESCALA SEMANAL DE EXAMINADORES — {html.escape(str(banca)).upper()}"
        f" — {html.escape(rotulo_semana).upper()} DE {html.escape(str(mes)).upper()}/{ano}",
    )

    dados = [cabecalho_colunas] + [
        [html.escape(str(celula)) for celula in linha] for linha in linhas_matriz
    ]
    largura_primeira = 150
    largura_demais = max(
        (LARGURA_UTIL_A4_PAISAGEM - largura_primeira)
        / max(len(cabecalho_colunas) - 1, 1),
        40,
    )
    tabela = Table(
        dados,
        repeatRows=1,
        colWidths=[largura_primeira]
        + [largura_demais] * (len(cabecalho_colunas) - 1),
    )
    # Linhas de resumo (total, efetivo, saldo, sobra) em destaque.
    resumo = [
        i
        for i, linha in enumerate(dados)
        if i and str(linha[0]).isupper()
    ]
    estilo = _estilo_tabela_padrao(resumo or [len(dados) - 1])
    estilo.add("ALIGN", (0, 1), (0, -1), "LEFT")
    tabela.setStyle(estilo)
    elementos.append(tabela)
    elementos.append(Spacer(1, 10))

    elementos.append(Paragraph("EQUIPES ITINERANTES NA SEMANA", estilos["grupo"]))
    if equipes:
        linhas_equipes = [
            ["Equipe", "Período", "Municípios", "Examinadores", "Observação"]
        ]
        for equipe in equipes:
            linhas_equipes.append(
                [
                    f"Banca {int(equipe['Numero']):02d}",
                    f"{formatar_br(equipe['Inicio Janela'])} a"
                    f" {formatar_br(equipe['Fim Janela'])}",
                    html.escape(" / ".join(equipe["Destinos"])),
                    f"{int(equipe['Examinadores']):02d}",
                    html.escape(equipe.get("Observacao", "") or "-"),
                ]
            )
        tabela_equipes = Table(
            linhas_equipes,
            repeatRows=1,
            colWidths=[60, 120, 240, 70, LARGURA_UTIL_A4_PAISAGEM - 490],
        )
        estilo_equipes = _estilo_tabela_padrao()
        estilo_equipes.add("ALIGN", (2, 1), (2, -1), "LEFT")
        estilo_equipes.add("ALIGN", (4, 1), (4, -1), "LEFT")
        tabela_equipes.setStyle(estilo_equipes)
        elementos.append(tabela_equipes)
    else:
        elementos.append(
            Paragraph("Nenhuma equipe em viagem nesta semana.", estilos["corpo"])
        )

    elementos.append(Spacer(1, 8))
    elementos.append(
        Paragraph(
            "Legenda: número = examinadores escalados · <b>Viagens</b> ="
            " examinadores fora da sede, contando o dia inteiro de deslocamento ·"
            " <b>x</b> = sem vagas lançadas · <b>-</b> = dia fora da grade da"
            " localidade · <b>LIVRE P/</b> = sobra para os postos montados por"
            f" último (mínimo {MINIMO_SOBRA_POSTOS}). Gerado em "
            f"{datetime.datetime.now().strftime('%d/%m/%Y às %H:%M')}.",
            estilos["nota"],
        )
    )
    documento.build(elementos)
    return buffer.getvalue()


# --- PDF: escala mensal de viagens itinerantes ------------------------------
def gerar_pdf_escala_viagens(
    viagens: Sequence[Viagem], mes: str, ano: int, mes_numero: int
) -> bytes:
    """Escala de viagens do mês, com preposto e veículo por equipe.

    A logística (01 preposto e 01 veículo por equipe itinerante) é atribuída
    automaticamente aqui: não há cadastro disso na interface, conforme
    solicitado — é informação que existe apenas neste documento.
    """
    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=18,
        leftMargin=18,
        topMargin=18,
        bottomMargin=18,
        title=f"Escala de viagens — {mes}/{ano}",
    )
    estilos = _estilos_pdf()
    elementos: list = []
    _cabecalho_oficial(
        elementos,
        estilos,
        "ESCALA DE VIAGENS DAS BANCAS ITINERANTES — "
        f"{html.escape(str(mes)).upper()} DE {ano}",
    )

    primeiro = datetime.date(ano, mes_numero, 1)
    ultimo = datetime.date(ano, mes_numero, calendar.monthrange(ano, mes_numero)[1])

    grupos = agrupar_trechos_por_equipe(viagens)
    do_mes = []
    for grupo in grupos:
        dias = [
            d
            for t in grupo["Trechos"]
            for d in dias_efetivos_da_viagem(t)
            if primeiro <= d <= ultimo
        ]
        if not dias:
            continue
        copia = dict(grupo)
        copia["Inicio Mes"] = min(dias)
        copia["Fim Mes"] = max(dias)
        do_mes.append(copia)

    if not do_mes:
        elementos.append(
            Paragraph(
                "<b>Nenhuma viagem programada para o mês selecionado.</b>",
                estilos["subtitulo"],
            )
        )
        documento.build(elementos)
        return buffer.getvalue()

    total_examinadores_mes = 0
    total_prepostos = 0
    total_veiculos = 0

    for banca in sorted({g["Banca"] for g in do_mes}):
        elementos.append(Paragraph(html.escape(banca).upper(), estilos["grupo"]))

        linhas = [
            [
                "Banca\nItinerante",
                "Período",
                "Municípios",
                "Equipe",
                "Veículo",
                "Detalhamento por dia",
            ]
        ]
        for grupo in sorted(
            (g for g in do_mes if g["Banca"] == banca), key=lambda g: g["Numero"]
        ):
            examinadores = int(grupo["Examinadores"])
            total_examinadores_mes += examinadores
            total_prepostos += PREPOSTOS_POR_EQUIPE
            total_veiculos += VEICULOS_POR_EQUIPE

            detalhes = []
            for trecho in grupo["Trechos"]:
                dias = [
                    d
                    for d in dias_efetivos_da_viagem(trecho)
                    if primeiro <= d <= ultimo
                ]
                if not dias:
                    continue
                turnos = {turno_da_viagem_no_dia(trecho, d) for d in dias}
                if turnos == {TURNO_INTEGRAL}:
                    sufixo = ""
                else:
                    sufixo = " (" + ", ".join(
                        f"{formatar_curto(d)} {turno_da_viagem_no_dia(trecho, d).lower()}"
                        for d in dias
                        if turno_da_viagem_no_dia(trecho, d) != TURNO_INTEGRAL
                    ) + ")"
                chegada, retorno = horarios_limite_da_viagem(trecho)
                limites = [
                    texto
                    for texto in (
                        f"provas a partir de {chegada} no 1º dia" if chegada else "",
                        f"até {retorno} no último dia" if retorno else "",
                    )
                    if texto
                ]
                detalhes.append(
                    f"{trecho.get('Destino')}: {formatar_curto(min(dias))} a"
                    f" {formatar_curto(max(dias))}{sufixo}"
                    + (f" [{', '.join(limites)}]" if limites else "")
                )

            observacoes = [
                str(t.get("Observações", "")).strip()
                for t in grupo["Trechos"]
                if str(t.get("Observações", "")).strip()
            ]
            if observacoes:
                detalhes.append("Obs.: " + "; ".join(dict.fromkeys(observacoes)))

            linhas.append(
                [
                    f"Banca {int(grupo['Numero']):02d}",
                    f"{formatar_br(grupo['Inicio Mes'])}\na"
                    f" {formatar_br(grupo['Fim Mes'])}",
                    Paragraph(
                        html.escape(" e ".join(grupo["Destinos"])), estilos["corpo"]
                    ),
                    f"{examinadores:02d} examinadores\n+"
                    f" {PREPOSTOS_POR_EQUIPE:02d} preposto",
                    f"{VEICULOS_POR_EQUIPE:02d} veículo",
                    Paragraph(html.escape(" · ".join(detalhes)), estilos["corpo"]),
                ]
            )

        tabela = Table(
            linhas,
            repeatRows=1,
            colWidths=[62, 92, 150, 96, 54, LARGURA_UTIL_A4_PAISAGEM - 454],
        )
        estilo = _estilo_tabela_padrao()
        estilo.add("ALIGN", (2, 1), (2, -1), "LEFT")
        estilo.add("ALIGN", (5, 1), (5, -1), "LEFT")
        estilo.add("VALIGN", (0, 1), (-1, -1), "TOP")
        tabela.setStyle(estilo)
        elementos.append(tabela)
        elementos.append(Spacer(1, 8))

    resumo = Table(
        [
            ["TOTAL DO MÊS", "Equipes", "Examinadores", "Prepostos", "Veículos"],
            [
                "",
                f"{len(do_mes):02d}",
                f"{total_examinadores_mes:02d}",
                f"{total_prepostos:02d}",
                f"{total_veiculos:02d}",
            ],
        ],
        colWidths=[140, 90, 110, 90, 90],
    )
    resumo.setStyle(_estilo_tabela_padrao([1]))
    elementos.append(resumo)
    elementos.append(Spacer(1, 6))
    elementos.append(
        Paragraph(
            "Cada banca itinerante conta com 01 preposto e 01 veículo atrelados à"
            " equipe de examinadores. Documento gerado em "
            f"{datetime.datetime.now().strftime('%d/%m/%Y às %H:%M')}.",
            estilos["nota"],
        )
    )
    documento.build(elementos)
    return buffer.getvalue()


# --- Excel de publicação ----------------------------------------------------
COL_DETALHE_PCD = "_detalhe_pcd"


def detalhe_pcd(df_vagas: pd.DataFrame) -> str:
    """Dias da semana e turnos em que cada coluna PCD teve vaga (JSON, para a
    observação da planilha). Texto e não lista: o DataFrame do resumo passa
    pelo cache do Streamlit, que precisa conseguir calcular o hash dele."""
    detalhes = {}
    for coluna in COLS_PCD:
        if coluna not in df_vagas.columns:
            continue
        com_vaga = df_vagas[df_vagas[coluna].fillna(0) > 0]
        if com_vaga.empty:
            continue
        detalhes[coluna] = {
            "dias": sorted(
                {str(d) for d in com_vaga["Dia da Semana"]},
                key=lambda d: DIAS_SEMANA_OPCOES.index(d) if d in DIAS_SEMANA_OPCOES else 9,
            ),
            "turnos": sorted({periodo_do_horario(h) for h in com_vaga["Horário"]}),
        }
    return _dump(detalhes)


@st.cache_data(show_spinner=False)
def gerar_excel_publicacao(df_resumo: pd.DataFrame, mes: str, ano: int | str) -> bytes:
    """Planilha no layout usado para publicação do calendário mensal.

    Ordem: São Luís, Imperatriz, Timon, Bacabal, Caxias e Santa Inês (o
    ``df_resumo`` já chega ordenado por ``chave_ordem_publicacao``). As vagas
    PCD somam na categoria de ampla concorrência e a célula ganha asterisco;
    a observação de cada asterisco sai no fim da planilha.
    """
    planilha = openpyxl.Workbook()
    aba = planilha.active
    aba.title = "Plan1"

    fonte = Font(name="Arial", size=10)
    centro = Alignment(horizontal="center", vertical="center")
    centro_quebra = Alignment(horizontal="center", vertical="center", wrap_text=True)
    esquerda = Alignment(horizontal="left", vertical="center")
    justificado = Alignment(horizontal="justify", vertical="top", wrap_text=True)
    fina = Side(border_style="thin", color="000000")
    borda = Border(left=fina, right=fina, top=fina, bottom=fina)

    colunas_vagas = list(COLS_CATEGORIA)
    larguras = {"A": 3, "B": 30.29}
    for indice in range(len(colunas_vagas)):
        larguras[chr(ord("C") + indice)] = 8.43
    larguras[chr(ord("C") + len(colunas_vagas))] = 25.29
    for letra, largura in larguras.items():
        aba.column_dimensions[letra].width = largura

    ultima_coluna = 2 + len(colunas_vagas) + 1
    linha_atual = 3
    observacoes: list[str] = []

    for banca, grupo in df_resumo.groupby("Banca", sort=False):
        nome_publicacao = str(banca).replace("Banca ", "")
        aba.merge_cells(
            start_row=linha_atual,
            start_column=2,
            end_row=linha_atual,
            end_column=ultima_coluna,
        )
        titulo = aba.cell(
            row=linha_atual,
            column=2,
            value=(
                f"Quantidade de Exames Mensais de {nome_publicacao} e Região"
                f" - {str(mes).upper()} DE {ano}"
            ),
        )
        titulo.font = fonte
        titulo.alignment = centro
        for coluna in range(2, ultima_coluna + 1):
            aba.cell(row=linha_atual, column=coluna).border = borda
        linha_atual += 1

        for indice, texto in enumerate(
            ["Cidade"] + colunas_vagas + ["Datas dos exames"], 2
        ):
            celula = aba.cell(row=linha_atual, column=indice, value=texto)
            celula.font = fonte
            celula.alignment = centro
            celula.border = borda
        linha_atual += 1

        for _, registro in grupo.iterrows():
            celula_cidade = aba.cell(
                row=linha_atual, column=2, value=str(registro["Localidade"])
            )
            celula_cidade.font = fonte
            celula_cidade.alignment = esquerda
            celula_cidade.border = borda

            try:
                detalhes = json.loads(str(registro.get(COL_DETALHE_PCD) or "{}"))
            except json.JSONDecodeError:
                detalhes = {}
            for indice, coluna in enumerate(colunas_vagas, 3):
                categoria = categoria_da_coluna(coluna)
                pcd = _num(registro.get(f"PCD {categoria}", 0))
                valor = _num(registro.get(coluna, 0)) + pcd
                celula = aba.cell(
                    row=linha_atual, column=indice, value=valor if valor > 0 else None
                )
                celula.font = fonte
                celula.alignment = centro
                celula.border = borda
                celula.number_format = "#,##0"
                if pcd > 0:
                    marcador = "*" * (len(observacoes) + 1)
                    celula.number_format = f'#,##0"{marcador}"'
                    info = detalhes.get(f"PCD {categoria}") or {}
                    observacoes.append(
                        observacao_pcd(
                            marcador,
                            categoria,
                            str(registro["Localidade"]),
                            pcd,
                            info.get("dias") or [],
                            info.get("turnos") or [],
                        )
                    )

            celula_datas = aba.cell(
                row=linha_atual,
                column=ultima_coluna,
                value=str(registro.get("Datas dos exames", "-")),
            )
            celula_datas.font = fonte
            celula_datas.alignment = centro_quebra
            celula_datas.number_format = "@"
            celula_datas.border = borda
            linha_atual += 1

        linha_atual += 1

    ultima_linha = max(linha_atual - 2, 3)
    for texto in observacoes:
        aba.merge_cells(
            start_row=linha_atual, start_column=2, end_row=linha_atual, end_column=ultima_coluna
        )
        celula = aba.cell(row=linha_atual, column=2, value=texto)
        celula.font = fonte
        celula.alignment = justificado
        aba.row_dimensions[linha_atual].height = 15 * max(2, math.ceil(len(texto) / 95))
        ultima_linha = linha_atual
        linha_atual += 1

    aba.print_area = f"A3:{chr(ord('A') + ultima_coluna - 1)}{ultima_linha}"
    aba.page_setup.orientation = "landscape"
    aba.page_setup.paperSize = aba.PAPERSIZE_A4
    aba.page_margins.left = 0.51
    aba.page_margins.right = 0.51
    aba.page_margins.top = 0.79
    aba.page_margins.bottom = 0.79
    aba.page_margins.header = 0.31
    aba.page_margins.footer = 0.31

    buffer = io.BytesIO()
    planilha.save(buffer)
    return buffer.getvalue()


# ===========================================================================
# 5. COMPONENTES DE INTERFACE
# ===========================================================================


def _botao_abrir_pdf(pdf_bytes: bytes, chave: str, rotulo: str) -> None:
    """Botão que abre o PDF em uma nova aba do navegador.

    Usa Blob + URL.createObjectURL em vez de um link ``data:``: o Chrome
    bloqueia navegação de nível superior para URLs ``data:``, então o link
    simples não abriria a aba.
    """
    b64 = base64.b64encode(pdf_bytes).decode("ascii")
    identificador = re.sub(r"[^A-Za-z0-9_]", "_", chave)
    componente = f"""
    <button id="abrir_{identificador}" style="
        width: 100%;
        padding: 0.55rem 0.75rem;
        border-radius: 6px;
        border: 1px solid #2B6CB0;
        background: #2B6CB0;
        color: #fff;
        font-weight: 600;
        font-size: 0.9rem;
        font-family: 'Source Sans Pro', sans-serif;
        cursor: pointer;">{html.escape(rotulo)}</button>
    <script>
      (function() {{
        const dados = "{b64}";
        const botao = document.getElementById("abrir_{identificador}");
        botao.addEventListener("click", function() {{
          const binario = atob(dados);
          const bytes = new Uint8Array(binario.length);
          for (let i = 0; i < binario.length; i++) {{
            bytes[i] = binario.charCodeAt(i);
          }}
          const blob = new Blob([bytes], {{ type: "application/pdf" }});
          const url = URL.createObjectURL(blob);
          const aba = window.open(url, "_blank");
          if (!aba) {{
            alert("O navegador bloqueou a nova aba. Libere os pop-ups deste site.");
          }}
          setTimeout(function() {{ URL.revokeObjectURL(url); }}, 60000);
        }});
      }})();
    </script>
    """
    components.html(componente, height=52)


def _previa_pdf(pdf_bytes: bytes, chave: str, altura: int = 620) -> None:
    """Pré-visualização embutida do PDF, para conferir sem sair da página."""
    b64 = base64.b64encode(pdf_bytes).decode("ascii")
    identificador = re.sub(r"[^A-Za-z0-9_]", "_", chave)
    componente = f"""
    <div id="previa_{identificador}" style="width:100%;height:{altura}px;"></div>
    <script>
      (function() {{
        const dados = "{b64}";
        const binario = atob(dados);
        const bytes = new Uint8Array(binario.length);
        for (let i = 0; i < binario.length; i++) {{
          bytes[i] = binario.charCodeAt(i);
        }}
        const blob = new Blob([bytes], {{ type: "application/pdf" }});
        const url = URL.createObjectURL(blob);
        const quadro = document.createElement("iframe");
        quadro.src = url;
        quadro.style.width = "100%";
        quadro.style.height = "100%";
        quadro.style.border = "1px solid #CBD5E0";
        quadro.style.borderRadius = "6px";
        document.getElementById("previa_{identificador}").appendChild(quadro);
      }})();
    </script>
    """
    components.html(componente, height=altura + 10)


def assinatura_dados(*partes: Any) -> str:
    """Assinatura estável de dados (inclusive DataFrames) para cache/PDF."""

    def normalizar(valor: Any) -> Any:
        if isinstance(valor, pd.DataFrame):
            return valor.to_dict(orient="records")
        if isinstance(valor, dict):
            return {str(k): normalizar(v) for k, v in valor.items()}
        if isinstance(valor, (list, tuple)):
            return [normalizar(v) for v in valor]
        return valor

    return _assinatura(normalizar(list(partes)))


def cabecalho_pagina(titulo: str, descricao: str = "") -> None:
    """Título padrão das páginas."""
    st.markdown(
        f'<div class="titulo-pagina"><div class="nome">{html.escape(titulo)}</div>'
        + (f'<div class="descricao">{html.escape(descricao)}</div>' if descricao else "")
        + "</div>",
        unsafe_allow_html=True,
    )


def _flash(tipo: str, mensagem: str) -> None:
    """Mensagem que sobrevive ao ``st.rerun()`` (antes, avisos como o de
    feriado na viagem sumiam porque o rerun vinha logo depois)."""
    st.session_state.setdefault("_flash", []).append((tipo, mensagem))


def mostrar_flash() -> None:
    for tipo, mensagem in st.session_state.pop("_flash", []):
        getattr(st, tipo, st.info)(mensagem)


def bloco_pdf(
    *,
    chave: str,
    rotulo_gerar: str,
    gerador,
    nome_arquivo: str,
    assinatura: str,
    ajuda: str | None = None,
    altura_previa: int = 620,
) -> None:
    """Fluxo padrão de PDF: gerar, conferir e só então baixar.

    O PDF fica guardado junto com a assinatura dos dados usados. Se os dados
    mudarem depois, o PDF guardado é descartado: antes dava para baixar uma
    versão antiga achando que era a atual.
    """
    chave_bytes = f"_pdf_{chave}"
    chave_previa = f"_previa_{chave}"

    guardado = st.session_state.get(chave_bytes)
    if guardado and guardado[0] != assinatura:
        st.session_state.pop(chave_bytes, None)
        st.caption("Os dados mudaram desde a última geração. Gere o PDF de novo.")

    if st.button(rotulo_gerar, key=f"btn_gerar_{chave}", help=ajuda):
        try:
            st.session_state[chave_bytes] = (assinatura, gerador())
            st.session_state[chave_previa] = True
        except Exception as erro:
            LOG.exception("Falha ao gerar o PDF %s", chave)
            st.error(f"Não foi possível gerar o PDF: {erro}")

    guardado = st.session_state.get(chave_bytes)
    if not guardado:
        return
    pdf_bytes = guardado[1]

    tamanho_kb = len(pdf_bytes) / 1024
    st.caption(f"Documento pronto ({tamanho_kb:.0f} KB). Confira antes de baixar.")

    col_abrir, col_baixar = st.columns(2)
    with col_abrir:
        _botao_abrir_pdf(pdf_bytes, chave, "🔎 Abrir em nova aba")
    with col_baixar:
        st.download_button(
            "⬇️ Baixar PDF",
            data=pdf_bytes,
            file_name=nome_arquivo,
            mime="application/pdf",
            key=f"btn_baixar_{chave}",
            **LARGURA_TOTAL,
        )

    if st.checkbox(
        "👁️ Pré-visualizar aqui na página",
        value=bool(st.session_state.get(chave_previa)),
        key=f"chk_previa_{chave}",
    ):
        _previa_pdf(pdf_bytes, chave, altura_previa)


# --- Acesso aos parâmetros guardados ----------------------------------------
def horarios_ativos() -> list[str]:
    inativos = set(st.session_state["horarios_inativos"])
    return sorted(
        (h for h in st.session_state["lista_horarios"] if h not in inativos),
        key=lambda h: _hora_em_minutos(h) or 0,
    )


def horarios_fim_turno() -> list[str]:
    return list(st.session_state.get("horarios_fim_turno", []))


def parametros_do_local(banca: str, local: str) -> dict[str, Any]:
    return dict(st.session_state["parametros_local_dict"].get(chave_local(banca, local)) or {})


def salvar_parametro_do_local(banca: str, local: str, nome: str, valor: Any) -> None:
    parametros = parametros_do_local(banca, local)
    if nome in parametros and parametros[nome] == valor:
        return
    parametros[nome] = valor
    st.session_state["parametros_local_dict"][chave_local(banca, local)] = parametros
    salvar_parametros_local(st.session_state["parametros_local_dict"])


def pistas_do_local(banca: str, local: str) -> int | None:
    """Pistas de moto do local (None = sem limite físico informado)."""
    parametros = parametros_do_local(banca, local)
    if "pistas_moto" in parametros:
        valor = parametros["pistas_moto"]
        return None if valor is None else max(0, _inteiro(valor))
    return PISTAS_MOTO_PADRAO.get(chave_local(banca, local))


def janelas_do_local(banca: str, local: str) -> dict[str, tuple[str, str]]:
    janelas = parametros_do_local(banca, local).get("janelas") or {}
    return {
        dia: (str(v[0]), str(v[1]))
        for dia, v in janelas.items()
        if isinstance(v, (list, tuple)) and len(v) == 2
    }


def examinadores_padrao_do_local(banca: str, local: str) -> tuple[int, int]:
    parametros = parametros_do_local(banca, local)
    return (
        _inteiro(parametros.get("examinadores_manha"), 4),
        _inteiro(parametros.get("examinadores_tarde"), 4),
    )


def dias_permitidos_do_local(banca: str, local: str) -> list[str]:
    return list(
        st.session_state["dias_permitidos_dict"].get(chave_local(banca, local), DIAS_UTEIS_PADRAO)
    )


def feriados_do_local(banca: str, local: str, ano: int) -> set[datetime.date]:
    """Feriados nacionais/estaduais, pontos facultativos marcados e feriados
    municipais digitados para a localidade."""
    datas = set(
        feriados_estaduais(
            int(ano), facultativos=tuple(st.session_state.get("facultativos_bloqueados", []))
        )
    )
    texto = st.session_state["feriados_locais_dict"].get(chave_local(banca, local), "")
    datas |= datas_de_feriado_do_texto(texto, int(ano))
    return datas


def corte_cat_a_do_local(banca: str, local: str) -> bool:
    """A Cat A cai pela metade no fim de turno deste posto?"""
    parametros = parametros_do_local(banca, local)
    if "corte_cat_a" in parametros:
        return bool(parametros["corte_cat_a"])
    return chave_local(banca, local) in CORTE_CAT_A_FIM_TURNO_PADRAO


def usa_toda_sobra(banca: str, local: str) -> bool:
    """Posto de sobra que recebe toda a sobra do dia (e não um número fixo)."""
    parametros = parametros_do_local(banca, local)
    if "toda_sobra" in parametros:
        return bool(parametros["toda_sobra"])
    return any(normalizar_texto(local) == normalizar_texto(p) for p in POSTOS_TODA_SOBRA_PADRAO)


# --- Efetivo de examinadores da banca ---------------------------------------
def chave_efetivo_dia(banca: str, data: datetime.date) -> str:
    return SEPARADOR_CHAVE.join([_sanitizar(banca), data.isoformat()])


def efetivo_padrao_da_banca(banca: str) -> int:
    return int(st.session_state["capacidade_bancas"].get(banca, 0) or 0)


def efetivo_sem_ajuste_diario(banca: str, data: datetime.date) -> int:
    """Efetivo do dia antes do ajuste por data: a disponibilidade semanal
    antiga (versão 4), se houver, ou o efetivo padrão da banca."""
    semanal = st.session_state["disponibilidade_semanal"].get(chave_disponibilidade(banca, data))
    if semanal is not None:
        return _inteiro(semanal)
    return efetivo_padrao_da_banca(banca)


def efetivo_do_dia(banca: str, data: datetime.date) -> int:
    """Examinadores da banca disponíveis no dia (já descontadas férias,
    licenças etc. informadas na página Efetivo)."""
    ajuste = st.session_state["efetivo_dia_dict"].get(chave_efetivo_dia(banca, data))
    if ajuste is not None:
        return _inteiro(ajuste)
    return efetivo_sem_ajuste_diario(banca, data)


def efetivo_do_mes(banca: str, ano: int, mes_numero: int) -> dict[datetime.date, int]:
    return {
        data: efetivo_do_dia(banca, data)
        for data in dias_do_intervalo(
            datetime.date(int(ano), mes_numero, 1),
            datetime.date(int(ano), mes_numero, calendar.monthrange(int(ano), mes_numero)[1]),
        )
    }


def definir_efetivo_padrao(banca: str, valor: int) -> None:
    capacidades = st.session_state["capacidade_bancas"]
    if int(capacidades.get(banca, 0) or 0) == int(valor):
        return
    capacidades[banca] = int(valor)
    salvar_config("capacidade_bancas", capacidades)


def definir_efetivo_do_dia(banca: str, data: datetime.date, valor: int) -> None:
    """Grava o ajuste do dia; igual ao valor sem ajuste, o ajuste é removido."""
    ajustes = st.session_state["efetivo_dia_dict"]
    chave = chave_efetivo_dia(banca, data)
    if int(valor) == efetivo_sem_ajuste_diario(banca, data):
        ajustes.pop(chave, None)
    else:
        ajustes[chave] = int(valor)


# --- Opções do gerador automático --------------------------------------------
def opcoes_do_gerador(chave: str) -> dict[str, Any]:
    """Demanda e opções guardadas do gerador para um calendário (chave do
    histórico): {"A": .., "PCD B": .., "periodo": [ini, fim], "dias": {..}}."""
    guardado = st.session_state["demandas_dict"].get(chave) or {}
    return guardado if isinstance(guardado, dict) else {}


# --- Grade do calendário ----------------------------------------------------
COL_VIAGEM = "_viagem"
COLUNAS_GRADE = (
    ["Data", "Dia da Semana", "Status", "Exam. M", "Exam. T", "Horário"]
    + COLS_VAGAS
    + ["Total", COL_AUTO]
)


def montar_grade_base(
    *,
    banca: str,
    local: str,
    ano: int,
    mes_numero: int,
    dias_permitidos: Sequence[str],
    horarios: Sequence[str],
    janelas: dict[str, tuple[str, str]],
    feriados: set[datetime.date],
    padrao_manha: int,
    padrao_tarde: int,
    viagens: Sequence[Viagem],
    bancas_config: dict[str, list[str]],
) -> tuple[pd.DataFrame, list[str]]:
    """Uma linha por (dia de atendimento x horário da janela do dia).

    Posto fixo (região metropolitana ou sede): dias da semana e janela de
    horários da localidade. Município itinerante: só os dias da viagem, e só
    os horários em que a equipe está lá (turno do dia, chegada no primeiro
    dia e retorno no último). Sem viagem no mês, não há grade.

    O status que o sistema impõe (feriado, horário fora da viagem, sede de
    banca regional com a banca viajando, dia fora da grade) vem marcado em
    ``Bloqueio Auto``, para poder ser desfeito quando a causa sumir.
    """
    fixo = eh_posto_fixo(bancas_config, banca, local)
    # A grade do município itinerante pertence à banca titular da viagem; a
    # banca de apoio entra só com examinadores (somados na grade da titular).
    viagens_do_local = [
        v
        for v in viagens
        if v.get("Destino") == local
        and (banca_vinculada(v, banca) if fixo else v.get("Banca") == banca)
    ]
    eh_regional = banca in BANCAS_REGIONAIS
    eh_sede = local == obter_sede(bancas_config, banca)

    total_dias = calendar.monthrange(int(ano), mes_numero)[1]
    registros: list[dict] = []
    datas_do_mes: list[str] = []

    for numero_dia in range(1, total_dias + 1):
        data = datetime.date(int(ano), mes_numero, numero_dia)
        nome_dia = DIAS_SEMANA_OPCOES[data.weekday()]
        viagens_no_dia = [v for v in viagens_do_local if viagem_cobre_dia(v, data)]
        if fixo:
            if nome_dia not in dias_permitidos and not viagens_no_dia:
                continue
            inicio, fim = janelas.get(nome_dia, JANELA_PADRAO)
            horarios_do_dia = horarios_na_janela(horarios, inicio, fim)
        else:
            if not viagens_no_dia:
                continue
            horarios_do_dia = list(horarios)
        if not horarios_do_dia:
            continue
        data_texto = data.strftime("%d/%m/%Y")
        datas_do_mes.append(data_texto)
        eh_feriado = data in feriados

        for horario in horarios_do_dia:
            periodo = periodo_do_horario(horario)
            viagem_no_horario = next(
                (v for v in viagens_no_dia if viagem_cobre_horario(v, data, horario)),
                None,
            )
            exam_manha = exam_tarde = 0
            if eh_feriado:
                status = STATUS_FERIADO
            elif viagem_no_horario:
                status = STATUS_DISPONIVEL
                efetivo = (
                    total_examinadores(viagem_no_horario)
                    if viagem_no_horario.get("Banca") == banca
                    else examinadores_da_banca_na_viagem(viagem_no_horario, banca)
                )
                exam_manha = efetivo if periodo == TURNO_MANHA else 0
                exam_tarde = efetivo if periodo == TURNO_TARDE else 0
            elif viagens_no_dia:
                # Equipe no município, mas em outro turno ou ainda na estrada.
                status = STATUS_INDISPONIVEL
            elif eh_regional and eh_sede and viagens_fora_da_sede(viagens, banca, data):
                status = STATUS_INDISPONIVEL
            elif nome_dia not in dias_permitidos:
                status = STATUS_INDISPONIVEL
            else:
                status = STATUS_DISPONIVEL
                exam_manha = padrao_manha if periodo == TURNO_MANHA else 0
                exam_tarde = padrao_tarde if periodo == TURNO_TARDE else 0

            registro = {
                "Data": data_texto,
                "Dia da Semana": nome_dia,
                "Status": status,
                "Exam. M": exam_manha,
                "Exam. T": exam_tarde,
                "Horário": horario,
            }
            for coluna in COLS_VAGAS:
                registro[coluna] = 0
            registro["Total"] = 0
            registro[COL_AUTO] = 1 if status in STATUS_BLOQUEADOS else 0
            registro[COL_VIAGEM] = 1 if viagem_no_horario and not eh_feriado else 0
            registros.append(registro)

    return pd.DataFrame(registros, columns=COLUNAS_GRADE + [COL_VIAGEM]), datas_do_mes


def mesclar_com_historico(df_base: pd.DataFrame, df_antigo: pd.DataFrame | None) -> pd.DataFrame:
    """Junta a grade atual com o que já foi lançado, sem perder lançamentos.

    * Bloqueio do sistema (feriado, viagem etc.): a linha fica bloqueada, mas
      as vagas lançadas antes continuam guardadas. Antes eram zeradas e se
      perdiam para sempre.
    * Bloqueio do sistema que deixou de existir: a linha volta a Disponível
      com as vagas de antes. Antes ela ficava presa como "Feriado".
    * Bloqueio escolhido pelo operador: é respeitado.
    * Linha de viagem: os examinadores vêm sempre da viagem (se o efetivo da
      viagem mudar, a grade acompanha).
    """
    if df_base.empty:
        return df_base.drop(columns=[COL_VIAGEM], errors="ignore").copy()
    indexado: dict[tuple[str, str], dict] = {}
    if df_antigo is not None and not df_antigo.empty:
        for antigo in df_antigo.to_dict(orient="records"):
            indexado[(str(antigo.get("Data")), str(antigo.get("Horário")))] = antigo

    linhas = []
    for base in df_base.to_dict(orient="records"):
        registro = dict(base)
        chave_linha = (base["Data"], base["Horário"])
        antigo = indexado.get(chave_linha)
        if antigo is not None:
            for coluna in COLS_VAGAS:
                registro[coluna] = _inteiro(antigo.get(coluna, 0))
            status_antigo = antigo.get("Status")
            bloqueio_auto_antigo = _inteiro(antigo.get(COL_AUTO))
            if base[COL_AUTO]:
                registro["Status"] = base["Status"]
                registro[COL_AUTO] = 1
            elif bloqueio_auto_antigo or status_antigo not in OPCOES_STATUS:
                registro["Status"] = base["Status"]
                registro[COL_AUTO] = 0
            else:
                registro["Status"] = status_antigo
                registro["Exam. M"] = _inteiro(antigo.get("Exam. M", 0))
                registro["Exam. T"] = _inteiro(antigo.get("Exam. T", 0))
                registro[COL_AUTO] = 0
            if base.get(COL_VIAGEM):
                registro["Exam. M"] = base["Exam. M"]
                registro["Exam. T"] = base["Exam. T"]
        registro["Total"] = (
            0
            if registro["Status"] in STATUS_BLOQUEADOS
            else sum(max(0, _inteiro(registro.get(c, 0))) for c in COLS_VAGAS)
        )
        registro.pop(COL_VIAGEM, None)
        linhas.append(registro)
    return pd.DataFrame(linhas, columns=COLUNAS_GRADE)


def grade_do_local(
    banca: str,
    local: str,
    ano: int,
    mes_numero: int,
    *,
    padrao_manha: int | None = None,
    padrao_tarde: int | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """(grade base, grade mesclada com o histórico, datas de atendimento),
    usando os parâmetros guardados da localidade."""
    if padrao_manha is None or padrao_tarde is None:
        padrao_manha, padrao_tarde = examinadores_padrao_do_local(banca, local)
    df_base, datas = montar_grade_base(
        banca=banca,
        local=local,
        ano=ano,
        mes_numero=mes_numero,
        dias_permitidos=dias_permitidos_do_local(banca, local),
        horarios=horarios_ativos(),
        janelas=janelas_do_local(banca, local),
        feriados=feriados_do_local(banca, local, ano),
        padrao_manha=padrao_manha,
        padrao_tarde=padrao_tarde,
        viagens=st.session_state["viagens_registradas"],
        bancas_config=st.session_state["bancas_config"],
    )
    chave = chave_historico(banca, MESES_LISTA[mes_numero - 1], ano, local)
    df_mesclado = mesclar_com_historico(
        df_base, st.session_state["historico_localidades"].get(chave)
    )
    return df_base, df_mesclado, datas


def dados_da_banca_no_mes(banca: str, mes: str, ano: int) -> dict[str, pd.DataFrame]:
    """Calendários da banca no mês, já na grade vigente.

    É a única porta de entrada dos dados para quadro, monitor, dashboard,
    Excel e PDFs: linhas de horários desativados ou de dias removidos não
    entram em nenhum total, e bloqueios novos (feriado, viagem) já valem mesmo
    para localidades que ainda não foram reabertas.
    """
    resultado: dict[str, pd.DataFrame] = {}
    mes_numero = numero_do_mes(mes)
    locais_da_banca = set(st.session_state["bancas_config"].get(banca, []))
    for chave in st.session_state["historico_localidades"]:
        partes = desmontar_chave_historico(chave)
        if not partes:
            continue
        banca_chave, mes_chave, ano_chave, local = partes
        if (banca_chave, mes_chave, ano_chave) != (banca, mes, str(ano)):
            continue
        if local not in locais_da_banca:
            continue
        _, df_mesclado, _ = grade_do_local(banca, local, int(ano), mes_numero)
        resultado[local] = df_mesclado
    return resultado


def colunas_de_validacao(
    df: pd.DataFrame, pistas: int | None, corte_a: bool = False
) -> pd.DataFrame:
    """Acrescenta as colunas calculadas: examinadores necessários, situação
    e horário quebrado de C/D/E."""
    copia = df.copy()
    fim = horarios_fim_turno()
    registros = copia.to_dict(orient="records")
    copia["Exam. Necess."] = [
        round(examinadores_necessarios(r, fim, corte_a), 2)
        if r.get("Status") == STATUS_DISPONIVEL
        else 0.0
        for r in registros
    ]
    problemas = [validar_linha(r, fim, pistas, corte_a) for r in registros]
    copia["Situação"] = [
        (SITUACAO_ERRO + ": " + "; ".join(p)) if p else SITUACAO_OK for p in problemas
    ]
    copia["Horário Quebrado"] = [descricao_horarios_quebrados(r) for r in registros]
    return copia


def violacoes(df: pd.DataFrame, pistas: int | None, corte_a: bool = False) -> list[str]:
    fim = horarios_fim_turno()
    mensagens = []
    for registro in df.to_dict(orient="records"):
        for problema in validar_linha(registro, fim, pistas, corte_a):
            mensagens.append(f"{registro['Data']} {registro['Horário']}: {problema}")
    return mensagens


def violacoes_do_local(banca: str, local: str, df: pd.DataFrame) -> list[str]:
    return violacoes(df, pistas_do_local(banca, local), corte_cat_a_do_local(banca, local))


def violacoes_da_banca(banca: str, dados_por_local: dict[str, pd.DataFrame]) -> list[str]:
    mensagens = []
    for local, df in dados_por_local.items():
        mensagens.extend(f"{local}, {m}" for m in violacoes_do_local(banca, local, df))
    return mensagens


# --- Postos de sobra (montados por último) ----------------------------------
def efetivo_diario_da_banca(
    banca: str,
    mes: str,
    ano: int,
    dados: dict[str, pd.DataFrame] | None = None,
    viagens: Sequence[Viagem] | None = None,
) -> dict[datetime.date, dict[str, Any]]:
    """``efetivo_diario`` já com a separação posto fixo × itinerante."""
    if dados is None:
        dados = dados_da_banca_no_mes(banca, mes, ano)
    if viagens is None:
        viagens = st.session_state["viagens_registradas"]
    fixos = postos_fixos(st.session_state["bancas_config"], banca)
    return efetivo_diario(dados, viagens, banca, int(ano), numero_do_mes(mes), fixos)


def sobra_do_mes(
    banca: str,
    mes: str,
    ano: int,
    excluir: Iterable[str],
    dados: dict[str, pd.DataFrame] | None = None,
    viagens: Sequence[Viagem] | None = None,
) -> dict[datetime.date, tuple[int, int]]:
    """(manhã, tarde) de examinadores livres por dia, sem contar as
    localidades de ``excluir`` (o próprio posto de sobra, ou todos eles)."""
    excluir = set(excluir)
    if dados is None:
        dados = dados_da_banca_no_mes(banca, mes, ano)
    restantes = {l: df for l, df in dados.items() if l not in excluir}
    totais = efetivo_diario_da_banca(banca, mes, ano, restantes, viagens)
    mes_numero = numero_do_mes(mes)
    sobra = {}
    for data, efetivo in efetivo_do_mes(banca, int(ano), mes_numero).items():
        usado = totais.get(data, {"M": 0, "T": 0})
        sobra[data] = (max(0, efetivo - usado["M"]), max(0, efetivo - usado["T"]))
    return sobra


def dias_dos_postos_de_sobra(banca: str, ano: int, mes_numero: int) -> list[datetime.date]:
    """Dias do mês em que algum posto de sobra atende (dia da semana da grade
    e sem feriado)."""
    postos = postos_de_sobra(st.session_state["bancas_config"], banca)
    if not postos:
        return []
    dias_semana = set().union(*(dias_permitidos_do_local(banca, p) for p in postos))
    feriados = feriados_do_local(banca, postos[0], int(ano))
    return [
        d
        for d in dias_do_intervalo(
            datetime.date(int(ano), mes_numero, 1),
            datetime.date(int(ano), mes_numero, calendar.monthrange(int(ano), mes_numero)[1]),
        )
        if nome_dia_semana(d) in dias_semana and d not in feriados
    ]


def alertas_de_sobra(
    banca: str,
    mes: str,
    ano: int,
    dados: dict[str, pd.DataFrame] | None = None,
    viagens: Sequence[Viagem] | None = None,
) -> list[tuple[datetime.date, int]]:
    """Dias em que sobram menos de ``MINIMO_SOBRA_POSTOS`` examinadores para
    os postos de sobra (Pátio DETRAN e Castelinho; Imperatriz)."""
    postos = postos_de_sobra(st.session_state["bancas_config"], banca)
    if not postos:
        return []
    sobra = sobra_do_mes(banca, mes, ano, postos, dados, viagens)
    alertas = []
    for data in dias_dos_postos_de_sobra(banca, int(ano), numero_do_mes(mes)):
        livres = min(sobra.get(data, (0, 0)))
        if livres < MINIMO_SOBRA_POSTOS:
            alertas.append((data, livres))
    return alertas


def texto_alerta_sobra(banca: str, alertas: Sequence[tuple[datetime.date, int]]) -> str:
    postos = " e ".join(postos_de_sobra(st.session_state["bancas_config"], banca))
    dias = ", ".join(f"{formatar_curto(d)} ({n})" for d, n in alertas[:12])
    return (
        f"Sobram menos de {MINIMO_SOBRA_POSTOS} examinadores para {postos} em"
        f" {len(alertas)} dia(s): {dias}{' ...' if len(alertas) > 12 else ''}."
        " Esses postos são montados por último com o que sobra do efetivo e não"
        f" podem ficar com menos de {MINIMO_SOBRA_POSTOS}. Reduza viagens ou"
        " examinadores das outras localidades nesses dias."
    )


# --- Geração automática -------------------------------------------------------
def _periodo_das_opcoes(
    opcoes: dict[str, Any], ano: int, mes_numero: int
) -> tuple[datetime.date, datetime.date]:
    primeiro = datetime.date(int(ano), mes_numero, 1)
    ultimo = datetime.date(int(ano), mes_numero, calendar.monthrange(int(ano), mes_numero)[1])
    periodo = opcoes.get("periodo") or []
    if len(periodo) == 2:
        inicio, fim = para_data(periodo[0]), para_data(periodo[1])
        if inicio and fim and primeiro <= inicio <= fim <= ultimo:
            return inicio, fim
    return primeiro, ultimo


def gerar_calendario_do_local(
    banca: str,
    local: str,
    ano: int,
    mes_numero: int,
    opcoes: dict[str, Any],
    *,
    exam_manha: int,
    exam_tarde: int,
    pistas: int | None,
    examinadores_por_dia: dict[datetime.date, tuple[int, int]] | None = None,
) -> dict[str, Any] | None:
    """Roda o gerador para uma localidade e grava o resultado.

    ``opcoes`` é o que o painel do gerador guarda: demanda por categoria
    (inclusive "PCD B"), período e dias da semana de cada categoria.
    ``examinadores_por_dia`` substitui os examinadores padrão (postos de
    sobra). Linhas de viagem mantêm os examinadores da viagem. Devolve
    {"gerado", "deficit", "demanda"} ou None se a localidade não tem grade.
    """
    df_base, df_completo, _ = grade_do_local(
        banca, local, ano, mes_numero, padrao_manha=exam_manha, padrao_tarde=exam_tarde
    )
    if df_completo.empty:
        return None
    periodo = _periodo_das_opcoes(opcoes, ano, mes_numero)
    de_viagem = {
        (r["Data"], r["Horário"]) for r in df_base.to_dict(orient="records") if r.get(COL_VIAGEM)
    }
    registros = df_completo.to_dict(orient="records")
    for registro in registros:
        data = para_data(registro["Data"])
        if (
            registro["Status"] != STATUS_DISPONIVEL
            or (registro["Data"], registro["Horário"]) in de_viagem
            or not data
            or not (periodo[0] <= data <= periodo[1])
        ):
            continue
        manha = periodo_do_horario(registro["Horário"]) == TURNO_MANHA
        if examinadores_por_dia is not None:
            exam_m, exam_t = examinadores_por_dia.get(data, (0, 0))
        else:
            exam_m, exam_t = exam_manha, exam_tarde
        registro["Exam. M"] = exam_m if manha else 0
        registro["Exam. T"] = 0 if manha else exam_t

    demanda = {c: _inteiro(opcoes.get(c)) for c in CATEGORIAS}
    demanda[COL_PCD_GERADA] = _inteiro(opcoes.get(COL_PCD_GERADA))
    dias = opcoes.get("dias") or {}
    dias_por_categoria = {
        c: dias[c] for c in CATEGORIAS + [COL_PCD_GERADA] if isinstance(dias.get(c), list)
    }
    if COL_PCD_GERADA not in dias_por_categoria:
        dias_por_categoria[COL_PCD_GERADA] = list(DIAS_PCD_PADRAO)
    gerados, gerado, deficit = gerar_distribuicao(
        registros,
        demanda,
        horarios_fim_turno(),
        pistas,
        periodo=periodo,
        dias_por_categoria=dias_por_categoria,
        corte_a=corte_cat_a_do_local(banca, local),
    )
    chave = chave_historico(banca, MESES_LISTA[mes_numero - 1], ano, local)
    df_novo = _normalizar_grade(pd.DataFrame(gerados, columns=COLUNAS_GRADE))
    _guardar_grade(chave, df_novo)
    _renovar_editor(chave)
    return {"gerado": gerado, "deficit": deficit, "demanda": demanda}


def resumo_da_geracao(local: str, resultado: dict[str, Any]) -> list[tuple[str, str]]:
    """Mensagens (tipo, texto) depois de uma geração."""
    demanda, gerado, deficit = resultado["demanda"], resultado["gerado"], resultado["deficit"]
    chaves = [c for c in CATEGORIAS + [COL_PCD_GERADA] if demanda.get(c)]
    rotulo = lambda c: c if c == COL_PCD_GERADA else f"Cat {c}"  # noqa: E731
    mensagens = [
        (
            "success",
            f"Calendário gerado para {local}. Vagas distribuídas: "
            + (", ".join(f"{rotulo(c)}: {gerado.get(c, 0)}" for c in chaves) or "nenhuma")
            + ".",
        )
    ]
    if gerado.get("A", 0) > demanda.get("A", 0):
        mensagens.append(
            (
                "info",
                f"Cat A fechada em blocos de {VAGAS_CAT_A_POR_DUPLA} por pista: a demanda de"
                f" {demanda['A']} virou {gerado['A']} vagas.",
            )
        )
    faltas = {c: v for c, v in deficit.items() if v > 0}
    if faltas:
        mensagens.append(
            (
                "warning",
                "Capacidade insuficiente para toda a demanda. Faltaram: "
                + ", ".join(f"{rotulo(c)}: {v}" for c, v in faltas.items())
                + ". Aumente examinadores, dias, horários ou pistas e gere de novo.",
            )
        )
    return mensagens


def examinadores_do_posto_de_sobra(
    banca: str, local: str, mes: str, ano: int
) -> dict[datetime.date, tuple[int, int]]:
    """Examinadores por dia de um posto de sobra: toda a sobra do dia ou o
    número fixo da localidade, limitado à sobra."""
    sobra = sobra_do_mes(banca, mes, ano, [local])
    if usa_toda_sobra(banca, local):
        return sobra
    fixo_m, fixo_t = examinadores_padrao_do_local(banca, local)
    return {d: (min(fixo_m, m), min(fixo_t, t)) for d, (m, t) in sobra.items()}


def gerar_postos_de_sobra(banca: str, mes: str, ano: int) -> list[tuple[str, str]]:
    """Gera os postos de sobra da banca com o que sobrou do efetivo.

    Primeiro os de número fixo (ex.: Castelinho), depois o que recebe toda a
    sobra (Pátio DETRAN), cada um com a demanda e as opções guardadas no
    próprio gerador. Deve ser feito depois das demais localidades.
    """
    postos = postos_de_sobra(st.session_state["bancas_config"], banca)
    mensagens: list[tuple[str, str]] = []
    mes_numero = numero_do_mes(mes)
    for local in sorted(postos, key=lambda l: usa_toda_sobra(banca, l)):
        opcoes = opcoes_do_gerador(chave_historico(banca, mes, ano, local))
        if not any(_inteiro(opcoes.get(c)) for c in CATEGORIAS + [COL_PCD_GERADA]):
            mensagens.append(
                ("warning", f"{local}: sem demanda informada no gerador; posto não gerado.")
            )
            continue
        resultado = gerar_calendario_do_local(
            banca,
            local,
            int(ano),
            mes_numero,
            opcoes,
            exam_manha=examinadores_padrao_do_local(banca, local)[0],
            exam_tarde=examinadores_padrao_do_local(banca, local)[1],
            pistas=pistas_do_local(banca, local),
            examinadores_por_dia=examinadores_do_posto_de_sobra(banca, local, mes, ano),
        )
        if resultado is None:
            mensagens.append(("warning", f"{local}: nenhuma data de atendimento no mês."))
            continue
        mensagens.extend(resumo_da_geracao(local, resultado))
    return mensagens


# ===========================================================================
# 6. PÁGINAS
# ===========================================================================


def _seletor_mes_ano(prefixo: str, colunas=None) -> tuple[str, int, int]:
    """Mês e ano padrão (mês atual, ano atual, com o ano anterior disponível)."""
    if colunas is None:
        colunas = st.columns(2)
    mes_nome = colunas[0].selectbox(
        "Mês:", MESES_LISTA, index=indice_mes_atual(), key=f"{prefixo}_mes"
    )
    ano = colunas[1].selectbox(
        "Ano:", anos_disponiveis(), index=indice_ano_atual(), key=f"{prefixo}_ano"
    )
    return mes_nome, int(ano), numero_do_mes(mes_nome)


# --- Página: quadro matriz de examinadores -----------------------------------
def _agrupar_por_semana(
    dias: list[datetime.date],
) -> dict[datetime.date, list[datetime.date]]:
    """Agrupa dias úteis pela segunda-feira da semana (única, ao contrário do
    número ISO, que se repete na virada do ano)."""
    semanas: dict[datetime.date, list[datetime.date]] = {}
    for dia in dias:
        segunda = dia - datetime.timedelta(days=dia.weekday())
        semanas.setdefault(segunda, []).append(dia)
    return dict(sorted(semanas.items()))


def _celula_do_quadro(
    banca: str,
    local: str,
    dia: datetime.date,
    por_data: dict[str, list[dict]],
    viagens: Sequence[Viagem],
) -> str:
    """Conteúdo da célula do quadro matriz (mesma regra do efetivo diário).

    Localidade atendida por equipe itinerante mostra só a marcação de
    viagem: o efetivo dela entra uma única vez na linha de total, pela
    equipe, e não somado por município.
    """
    viagem = viagem_do_local_no_dia(viagens, banca, local, dia)
    if viagem:
        numero = int(viagem.get("Numero Banca Itinerante", 0) or 0)
        turno = turno_da_viagem_no_dia(viagem, dia)
        marcador = f"Viagem {numero:02d}"
        if turno in (TURNO_MANHA, TURNO_TARDE):
            marcador += " (M)" if turno == TURNO_MANHA else " (T)"
        return marcador

    if nome_dia_semana(dia) not in dias_permitidos_do_local(banca, local):
        return "-"
    if dia in feriados_do_local(banca, local, dia.year):
        return "Feriado"
    do_dia = por_data.get(dia.strftime("%d/%m/%Y"), [])
    if do_dia and all(r.get("Status") == STATUS_FERIADO for r in do_dia):
        return "Feriado"
    manha, tarde = efetivo_do_grupo(do_dia)
    if manha or tarde:
        return str(pico_diario(manha, tarde))
    return "x"


ROTULO_LINHA_VIAGENS = "Viagens (equipes itinerantes)"
ROTULO_LINHA_TOTAL = "TOTAL EM CAMPO"
ROTULO_LINHA_EFETIVO = "EFETIVO DISPONÍVEL"
ROTULO_LINHA_SALDO = "SALDO DO DIA"


def nome_curto(local: str) -> str:
    """'São Luís - Pátio DETRAN' → 'Pátio DETRAN'."""
    return local.split(" - ", 1)[-1]


def rotulo_linha_sobra(postos: Sequence[str]) -> str:
    return "LIVRE P/ " + " E ".join(nome_curto(p).upper() for p in postos)


def _montar_matriz_semana(
    banca: str,
    dias_semana: list[datetime.date],
    por_local: dict[str, dict[str, list[dict]]],
    totais: dict[datetime.date, dict[str, Any]],
    efetivos: dict[datetime.date, int],
    sobra_postos: dict[datetime.date, tuple[int, int]] | None = None,
) -> tuple[list[str], list[list[str]]]:
    """Quadro da semana: um posto fixo por linha e as viagens somadas.

    Nas bancas de São Luís e Imperatriz o quadro mostra só a região
    metropolitana (ou Imperatriz) e uma linha "Viagens" com o total de
    examinadores fora da sede no dia, incluindo os dias de deslocamento. O
    nome de cada município itinerante fica no detalhamento das equipes.
    """
    bancas_config = st.session_state["bancas_config"]
    viagens = st.session_state["viagens_registradas"]
    linhas: list[list[str]] = []
    for local in postos_fixos(bancas_config, banca):
        por_data = por_local.get(local, {})
        linhas.append(
            [local] + [_celula_do_quadro(banca, local, dia, por_data, viagens) for dia in dias_semana]
        )

    def valor(numero: int) -> str:
        return str(numero) if numero else "-"

    linhas.append(
        [ROTULO_LINHA_VIAGENS]
        + [valor(totais.get(d, {}).get("viagem", 0)) for d in dias_semana]
    )
    picos = {
        d: pico_diario(totais.get(d, {}).get("M", 0), totais.get(d, {}).get("T", 0))
        for d in dias_semana
    }
    linhas.append([ROTULO_LINHA_TOTAL] + [str(picos[d]) for d in dias_semana])
    linhas.append([ROTULO_LINHA_EFETIVO] + [str(efetivos.get(d, 0)) for d in dias_semana])
    linhas.append(
        [ROTULO_LINHA_SALDO] + [str(efetivos.get(d, 0) - picos[d]) for d in dias_semana]
    )
    if sobra_postos is not None:
        postos = postos_de_sobra(bancas_config, banca)
        linhas.append(
            [rotulo_linha_sobra(postos)]
            + [str(min(sobra_postos.get(d, (0, 0)))) for d in dias_semana]
        )

    cabecalho = ["Localidade"] + [
        f"{d.strftime('%d/%m')} ({nome_dia_semana(d)[:3]})" for d in dias_semana
    ]
    return cabecalho, linhas


def _estilizar_quadro(
    df: pd.DataFrame, dias_sobra: set[str], rotulo_sobra: str | None
):
    """Destaque das linhas de total e dos dias críticos do quadro."""
    resumo = {ROTULO_LINHA_TOTAL, ROTULO_LINHA_EFETIVO, ROTULO_LINHA_SALDO}

    def estilo_linha(linha: pd.Series) -> list[str]:
        rotulo = linha.iloc[0]
        estilos = []
        for coluna, celula in linha.items():
            css = ""
            if rotulo in resumo or rotulo == rotulo_sobra:
                css = f"background-color: {COR_CLARA}; font-weight: 700; color: {COR_PRIMARIA};"
            elif rotulo == ROTULO_LINHA_VIAGENS:
                css = f"background-color: #E6F4EA; color: {COR_VIAGEM}; font-weight: 600;"
            if coluna != "Localidade":
                texto = str(celula)
                negativo = texto.lstrip("-").isdigit() and texto.startswith("-") and texto != "-"
                if rotulo == ROTULO_LINHA_SALDO and negativo:
                    css = f"background-color: {COR_ERRO_BG}; color: {COR_ERRO_TXT}; font-weight: 700;"
                if rotulo == rotulo_sobra and coluna in dias_sobra:
                    css = f"background-color: {COR_ERRO_BG}; color: {COR_ERRO_TXT}; font-weight: 700;"
            estilos.append(css)
        return estilos

    return df.style.apply(estilo_linha, axis=1)


def _painel_fechamento(banca: str, mes_nome: str, ano: int) -> None:
    """Botão que gera os postos de sobra depois das demais localidades."""
    postos = postos_de_sobra(st.session_state["bancas_config"], banca)
    if not postos:
        return
    ordem = sorted(postos, key=lambda l: usa_toda_sobra(banca, l))
    with st.expander(
        f"Fechamento do mês: gerar {' e '.join(nome_curto(p) for p in ordem)} com a sobra",
        expanded=False,
    ):
        st.caption(
            "Estes postos são montados por último: recebem o que sobra do efetivo de cada"
            " dia depois das demais localidades e das viagens (inclusive deslocamento)."
            " Ordem: "
            + " → ".join(
                f"{nome_curto(p)} ({'toda a sobra' if usa_toda_sobra(banca, p) else 'número fixo, limitado à sobra'})"
                for p in ordem
            )
            + ". A demanda e as opções de cada posto vêm do gerador da página Montar Calendário."
        )
        if st.button(
            "Gerar postos de sobra",
            key=f"btn_sobra_{_assinatura(banca)[:8]}_{mes_nome}_{ano}",
            type="primary",
            icon=":material/auto_mode:",
        ):
            for tipo, mensagem in gerar_postos_de_sobra(banca, mes_nome, ano):
                _flash(tipo, mensagem)
            st.rerun()


def pagina_quadro() -> None:
    cabecalho_pagina(
        "Quadro de Examinadores",
        "Escala diária da região metropolitana, viagens e saldo do efetivo.",
    )

    bancas_config = st.session_state["bancas_config"]
    disponiveis = [b for b in BANCAS_COM_QUADRO_MATRIZ if b in bancas_config]
    if not disponiveis:
        st.info("Nenhuma das bancas com quadro matriz está cadastrada no momento.")
        return

    col_banca, col_mes, col_ano = st.columns([2, 1, 1])
    banca = col_banca.selectbox("Banca", disponiveis, key="q_banca")
    mes_nome, ano, mes_numero = _seletor_mes_ano("q", [col_mes, col_ano])

    if not bancas_config.get(banca):
        st.warning(f"A banca {banca} não possui localidades cadastradas.")
        return

    viagens = st.session_state["viagens_registradas"]
    dados = dados_da_banca_no_mes(banca, mes_nome, ano)
    totais = efetivo_diario_da_banca(banca, mes_nome, ano, dados, viagens)
    efetivos = efetivo_do_mes(banca, ano, mes_numero)
    postos_sobra = postos_de_sobra(bancas_config, banca)
    sobra = sobra_do_mes(banca, mes_nome, ano, postos_sobra, dados, viagens) if postos_sobra else None
    alertas = alertas_de_sobra(banca, mes_nome, ano, dados, viagens)
    por_local = {local: registros_por_data(df) for local, df in dados.items()}

    dias_mes = dias_do_intervalo(
        datetime.date(ano, mes_numero, 1),
        datetime.date(ano, mes_numero, calendar.monthrange(ano, mes_numero)[1]),
    )
    picos = {d: pico_diario(totais.get(d, {}).get("M", 0), totais.get(d, {}).get("T", 0)) for d in dias_mes}
    # Dias úteis, e sábados só quando há alguém em campo.
    dias_quadro = [d for d in dias_mes if d.weekday() < 5 or (d.weekday() == 5 and picos[d])]
    estouros = [d for d in dias_mes if picos[d] > efetivos[d]]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Efetivo padrão da banca", efetivo_padrao_da_banca(banca))
    k2.metric("Maior pico no mês", max(picos.values(), default=0))
    k3.metric(
        "Menor saldo diário",
        min((efetivos[d] - picos[d] for d in dias_quadro), default=0),
    )
    k4.metric(
        f"Dias com sobra < {MINIMO_SOBRA_POSTOS}" if postos_sobra else "Dias acima do efetivo",
        len(alertas) if postos_sobra else len(estouros),
    )

    if estouros:
        st.error(
            "Efetivo excedido em: "
            + ", ".join(f"{formatar_curto(d)} ({picos[d]}/{efetivos[d]})" for d in estouros)
            + ". Ajuste o efetivo na página Efetivo ou reduza a escala."
        )
    if alertas:
        st.error(texto_alerta_sobra(banca, alertas))
    sem_viagem = sorted({l for v in totais.values() for l in v.get("sem_viagem", [])})
    if sem_viagem:
        st.warning(
            "Municípios itinerantes com vagas lançadas fora de viagem cadastrada: "
            + ", ".join(sem_viagem)
            + ". Eles entram na linha Viagens; cadastre a viagem para manter o"
            " controle alinhado."
        )

    _painel_fechamento(banca, mes_nome, ano)

    rotulo_sobra = rotulo_linha_sobra(postos_sobra) if postos_sobra else None
    dias_alerta = {d for d, _ in alertas}
    for indice, (_segunda, dias_semana) in enumerate(
        _agrupar_por_semana(dias_quadro).items(), 1
    ):
        rotulo = (
            f"Semana {indice} ({dias_semana[0].strftime('%d')} a"
            f" {dias_semana[-1].strftime('%d')} de {mes_nome})"
        )
        st.markdown(f'<div class="titulo-secao">{html.escape(rotulo)}</div>', unsafe_allow_html=True)

        cabecalho, linhas = _montar_matriz_semana(
            banca, dias_semana, por_local, totais, efetivos, sobra
        )
        df_semana = pd.DataFrame(linhas, columns=cabecalho)
        colunas_alerta = {
            cabecalho[i + 1] for i, d in enumerate(dias_semana) if d in dias_alerta
        }
        st.dataframe(
            _estilizar_quadro(df_semana, colunas_alerta, rotulo_sobra),
            **LARGURA_TOTAL,
            hide_index=True,
            height=(len(df_semana) + 1) * 35 + 3,
        )

        equipes = equipes_ativas_no_intervalo(viagens, banca, dias_semana[0], dias_semana[-1])
        if equipes:
            total_em_viagem = sum(int(e["Examinadores"]) for e in equipes)
            with st.expander(
                f"Detalhe das viagens: {len(equipes)} equipe(s), {total_em_viagem} examinador(es)",
                expanded=False,
            ):
                for equipe in equipes:
                    st.markdown(
                        f"**Banca {int(equipe['Numero']):02d}** ({equipe['Banca']}) ·"
                        f" {' / '.join(equipe['Destinos'])} ·"
                        f" {int(equipe['Examinadores'])} examinador(es)"
                    )
                    for trecho in equipe["Trechos"]:
                        inicio_t, fim_t = para_data(trecho["Data Inicio"]), para_data(trecho["Data Fim"])
                        if fim_t < dias_semana[0] or inicio_t > dias_semana[-1]:
                            continue
                        chegada, retorno = horarios_limite_da_viagem(trecho)
                        detalhe_turnos = descricao_turnos(trecho)
                        st.caption(
                            f"{trecho.get('Destino')}: {formatar_br(inicio_t)} a {formatar_br(fim_t)},"
                            f" {trecho.get('Turno', TURNO_INTEGRAL).lower()}"
                            + (f" · provas a partir de {chegada} no 1º dia" if chegada else "")
                            + (f" · até {retorno} no último dia" if retorno else "")
                            + (f" · ajustes: {detalhe_turnos}" if detalhe_turnos else "")
                            + (f" · {trecho['Observações']}" if trecho.get("Observações") else "")
                        )
        else:
            st.caption("Nenhuma equipe itinerante nesta semana.")

        for equipe in equipes:
            equipe["Observacao"] = "; ".join(
                f"{t.get('Destino')}: {descricao_turnos(t)}"
                for t in equipe["Trechos"]
                if descricao_turnos(t)
            )

        bloco_pdf(
            chave=f"semana_{banca}_{mes_nome}_{ano}_{indice}",
            rotulo_gerar=f"Gerar PDF da semana {indice}",
            gerador=lambda cab=cabecalho, lin=linhas, eq=equipes, rot=rotulo: (
                gerar_pdf_semana(cab, lin, eq, banca, mes_nome, ano, rot)
            ),
            nome_arquivo=(
                f"Escala_Semana{indice}_{banca}_{mes_nome}_{ano}.pdf".replace(" ", "_")
            ),
            assinatura=assinatura_dados(cabecalho, linhas, [e["Destinos"] for e in equipes]),
            ajuda="Gera o resumo da escala desta semana para impressão.",
            altura_previa=520,
        )

    st.caption(
        "Legenda: número = examinadores escalados no posto (maior turno do dia) ·"
        " **Viagens** = examinadores fora da sede, contando o dia inteiro de"
        " deslocamento · **x** = sem vagas lançadas · **-** = dia fora da grade ·"
        f" **{rotulo_sobra or 'Livre'}** = examinadores que sobram para os postos"
        f" montados por último (mínimo {MINIMO_SOBRA_POSTOS})."
    )


# --- Página: montagem do calendário detalhado --------------------------------
# Colunas calculadas (somente leitura) exibidas no editor. O Streamlit só
# aplica a cor do Styler em colunas desabilitadas do data_editor, por isso a
# faixa colorida (verde, cinza ou vermelha) fica nestas colunas.
COLUNAS_CALCULADAS = ["Exam. Necess.", "Situação", "Horário Quebrado"]
_COLUNAS_DESTACAVEIS_EDITOR = ["Data", "Dia da Semana", "Horário", "Total"] + COLUNAS_CALCULADAS
COLUNAS_EDITAVEIS = ["Status", "Exam. M", "Exam. T"] + COLS_VAGAS


def _cor_linha_calendario(linha: pd.Series) -> str:
    """Vermelho: estouro de capacidade/regra. Cinza: bloqueada. Verde: com
    vagas. Sem cor: ainda vazia."""
    if str(linha.get("Situação", "")).startswith(SITUACAO_ERRO):
        return f"background-color: {COR_ERRO_BG}; color: {COR_ERRO_TXT}; font-weight: 700;"
    if linha.get("Status") in STATUS_BLOQUEADOS:
        return f"background-color: {COR_INATIVO_BG}; color: {COR_INATIVO_TXT};"
    if _num(linha.get("Total")) > 0:
        return (
            f"background-color: {COR_LANCADO_BG}; color: {COR_LANCADO_TXT};"
            " font-weight: 600;"
        )
    return ""


def _estilizar_grade_calendario(df: pd.DataFrame):
    if df.empty:
        return df

    def _aplicar(linha: pd.Series) -> list[str]:
        estilo = _cor_linha_calendario(linha)
        return [estilo if col in _COLUNAS_DESTACAVEIS_EDITOR else "" for col in df.columns]

    return df.style.apply(_aplicar, axis=1)


def _normalizar_grade(df: pd.DataFrame) -> pd.DataFrame:
    """Tipos numéricos, status válido e Total recalculado."""
    copia = df.copy()
    for coluna in ["Exam. M", "Exam. T"] + COLS_VAGAS + [COL_AUTO]:
        if coluna in copia.columns:
            copia[coluna] = (
                pd.to_numeric(copia[coluna], errors="coerce").fillna(0).clip(lower=0).astype(int)
            )
    copia["Status"] = copia["Status"].where(copia["Status"].isin(OPCOES_STATUS), STATUS_DISPONIVEL)
    bloqueadas = copia["Status"].isin(STATUS_BLOQUEADOS)
    copia["Total"] = copia[COLS_VAGAS].sum(axis=1).where(~bloqueadas, 0).astype(int)
    return copia


def _aplicar_edicoes_pendentes(df: pd.DataFrame, chave_editor: str) -> pd.DataFrame:
    """Aplica as edições que o data_editor guarda no estado antes de desenhar.

    Assim a cor vermelha e as colunas calculadas já refletem a última edição
    no mesmo instante (validação em tempo real). As edições do Streamlit são
    valores absolutos por posição, então reaplicá-las é seguro.
    """
    estado = st.session_state.get(chave_editor)
    if not isinstance(estado, dict):
        return df
    copia = df.copy()
    for posicao, mudancas in (estado.get("edited_rows") or {}).items():
        try:
            indice = copia.index[int(posicao)]
        except (ValueError, IndexError):
            continue
        for coluna, valor in (mudancas or {}).items():
            if coluna in COLUNAS_EDITAVEIS:
                copia.at[indice, coluna] = valor
    return copia


def _guardar_grade(chave: str, df_completo: pd.DataFrame) -> None:
    """Atualiza a memória sem descartar linhas fora da grade atual e grava."""
    historico = st.session_state["historico_localidades"]
    antigo = historico.get(chave)
    if antigo is None or antigo.empty:
        historico[chave] = df_completo.copy()
    else:
        historico[chave] = (
            pd.concat([antigo, df_completo], ignore_index=True)
            .drop_duplicates(subset=["Data", "Horário"], keep="last")
            .reset_index(drop=True)
        )
    salvar_historico(chave, df_completo)


def _versao_editor(chave: str) -> int:
    return int(st.session_state.get(f"_ver_editor_{chave}", 0))


def _renovar_editor(chave: str) -> None:
    """Descarta as edições guardadas no widget (depois de uma geração
    automática, por exemplo, para elas não sobrescreverem o resultado)."""
    st.session_state[f"_ver_editor_{chave}"] = _versao_editor(chave) + 1


def _barra_lateral_calendario() -> dict[str, Any] | None:
    bancas_config = st.session_state["bancas_config"]
    if not bancas_config:
        st.warning("Cadastre uma banca na página de Gestão de Localidades.")
        return None

    sufixo = sufixo_widget()
    st.sidebar.header("Parâmetros do calendário")
    banca = st.sidebar.selectbox("Banca:", list(bancas_config.keys()), key="cal_banca")
    localidades = bancas_config.get(banca, [])
    if not localidades:
        st.warning(f"A banca **{banca}** ainda não possui localidades cadastradas.")
        return None
    fixos = postos_fixos(bancas_config, banca)
    local = st.sidebar.selectbox(
        "Local de atendimento:",
        localidades,
        key="cal_local",
        format_func=lambda l: l if l in fixos else f"{l} (itinerante)",
    )
    ano = int(
        st.sidebar.selectbox("Ano:", anos_disponiveis(), index=indice_ano_atual(), key="cal_ano")
    )
    mes_nome = st.sidebar.selectbox("Mês:", MESES_LISTA, index=indice_mes_atual(), key="cal_mes")
    chave_loc = chave_local(banca, local)
    id_local = _assinatura(chave_loc)[:10]
    itinerante = local not in fixos
    posto_sobra = local in postos_de_sobra(bancas_config, banca)

    st.sidebar.markdown("---")
    efetivo_atual = efetivo_padrao_da_banca(banca)
    novo_efetivo = int(
        st.sidebar.number_input(
            f"Efetivo diário de examinadores ({banca.replace('Banca ', '')})",
            min_value=0,
            max_value=300,
            value=efetivo_atual,
            step=1,
            # A chave leva o valor: se ele mudar em outra tela, o campo acompanha.
            key=f"cal_efetivo_{_assinatura(banca)[:8]}_{efetivo_atual}_{sufixo}",
            help=(
                "Examinadores da banca em um dia normal. Férias, licenças e outros"
                " ajustes por data ficam na página Efetivo."
            ),
        )
    )
    if novo_efetivo != efetivo_atual:
        definir_efetivo_padrao(banca, novo_efetivo)
        st.rerun()

    dias_permitidos = dias_permitidos_do_local(banca, local)
    exam_manha, exam_tarde = examinadores_padrao_do_local(banca, local)
    if itinerante:
        st.sidebar.info(
            "Município itinerante: as datas, os turnos, os horários de chegada e de"
            " retorno e os examinadores vêm da viagem cadastrada em Viagens"
            " Itinerantes."
        )
    else:
        st.sidebar.subheader("Dias de atendimento")
        dias_permitidos = st.sidebar.multiselect(
            "Dias de exame desta localidade:",
            options=DIAS_SEMANA_OPCOES,
            default=dias_permitidos,
            key=f"cal_dias_{id_local}_{sufixo}",
        )
        if dias_permitidos != dias_permitidos_do_local(banca, local):
            st.session_state["dias_permitidos_dict"][chave_loc] = dias_permitidos
            salvar_dias_permitidos(st.session_state["dias_permitidos_dict"])

        lista = sorted(st.session_state["lista_horarios"], key=lambda h: _hora_em_minutos(h) or 0)
        ativos = horarios_ativos()
        with st.sidebar.expander("Horário de atendimento por dia", expanded=False):
            st.caption("Primeiro e último horário de exame em cada dia.")
            if ativos and dias_permitidos:
                guardadas = janelas_do_local(banca, local)
                atuais = {
                    dia: list(guardadas.get(dia, (ativos[0], ativos[-1])))
                    for dia in dias_permitidos
                }
                df_janelas = pd.DataFrame(
                    [{"Dia": d, "Início": v[0], "Fim": v[1]} for d, v in atuais.items()]
                )
                editado = st.data_editor(
                    df_janelas,
                    hide_index=True,
                    num_rows="fixed",
                    column_config={
                        "Dia": st.column_config.TextColumn("Dia", disabled=True),
                        "Início": st.column_config.SelectboxColumn("Início", options=lista, required=True),
                        "Fim": st.column_config.SelectboxColumn("Fim", options=lista, required=True),
                    },
                    key=f"cal_janelas_{id_local}_{assinatura_dados(dias_permitidos, lista)[:8]}_{sufixo}",
                )
                novas = {
                    str(r["Dia"]): [str(r["Início"]), str(r["Fim"])]
                    for r in editado.to_dict(orient="records")
                }
                invertidas = [
                    d for d, (i, f) in novas.items()
                    if (_hora_em_minutos(i) or 0) > (_hora_em_minutos(f) or 0)
                ]
                if invertidas:
                    st.error("Início depois do fim em: " + ", ".join(invertidas))
                elif novas != atuais:
                    janelas_completas = {**{d: list(v) for d, v in guardadas.items()}, **novas}
                    salvar_parametro_do_local(banca, local, "janelas", janelas_completas)

        st.sidebar.subheader("Examinadores no posto")
        col_m, col_t = st.sidebar.columns(2)
        padrao_m, padrao_t = exam_manha, exam_tarde
        exam_manha = int(
            col_m.number_input("Manhã", 0, 100, padrao_m, key=f"cal_pad_m_{id_local}_{sufixo}")
        )
        exam_tarde = int(
            col_t.number_input("Tarde", 0, 100, padrao_t, key=f"cal_pad_t_{id_local}_{sufixo}")
        )
        if (exam_manha, exam_tarde) != (padrao_m, padrao_t):
            salvar_parametro_do_local(banca, local, "examinadores_manha", exam_manha)
            salvar_parametro_do_local(banca, local, "examinadores_tarde", exam_tarde)
        if posto_sobra:
            toda = st.sidebar.checkbox(
                "Usar toda a sobra do dia",
                value=usa_toda_sobra(banca, local),
                key=f"cal_toda_sobra_{id_local}_{sufixo}",
                help=(
                    "Posto de sobra: marcado, recebe todos os examinadores que sobram no"
                    " dia; desmarcado, recebe o número acima, limitado à sobra."
                ),
            )
            if toda != usa_toda_sobra(banca, local):
                salvar_parametro_do_local(banca, local, "toda_sobra", toda)
            st.sidebar.caption(
                "Posto montado por último, com a sobra do efetivo do dia"
                + (" (todos os examinadores livres)." if toda else " (número acima, limitado à sobra).")
            )
        else:
            st.sidebar.caption(
                "Valem para linhas novas e são aplicados a todas as linhas quando o"
                " gerador automático roda."
            )

    st.sidebar.subheader("Pistas de moto (Cat A)")
    pistas_atual = pistas_do_local(banca, local)
    sem_limite = st.sidebar.checkbox(
        "Sem limite físico informado",
        value=pistas_atual is None,
        key=f"cal_pistas_livre_{id_local}_{sufixo}",
    )
    if sem_limite:
        pistas = None
    else:
        pistas = int(
            st.sidebar.number_input(
                "Pistas ativas simultâneas:",
                min_value=0,
                max_value=10,
                value=pistas_atual if pistas_atual is not None else 1,
                key=f"cal_pistas_{id_local}_{sufixo}",
                help=f"1 pista comporta 1 dupla = {VAGAS_CAT_A_POR_DUPLA} exames de Cat A por horário.",
            )
        )
    if pistas != pistas_atual:
        salvar_parametro_do_local(banca, local, "pistas_moto", pistas)
    corte_atual = corte_cat_a_do_local(banca, local)
    corte_a = st.sidebar.checkbox(
        "Cat A pela metade no fim de turno (11:30)",
        value=corte_atual,
        key=f"cal_corte_a_{id_local}_{sufixo}",
        help="No Pátio do DETRAN, o horário de 11:30 oferta metade: 24, 24, 24 e 12 com 2 pistas.",
    )
    if corte_a != corte_atual:
        salvar_parametro_do_local(banca, local, "corte_cat_a", corte_a)

    st.sidebar.subheader("Feriados municipais / locais")
    texto_feriados = st.sidebar.text_area(
        "Um por linha (Ex: DD/MM - Motivo):",
        value=st.session_state["feriados_locais_dict"].get(chave_loc, ""),
        help="Formatos aceitos: DD/MM, DD/MM/AAAA ou AAAA-MM-DD",
        key=f"cal_feriados_{id_local}_{sufixo}",
    )
    if texto_feriados != st.session_state["feriados_locais_dict"].get(chave_loc, ""):
        st.session_state["feriados_locais_dict"][chave_loc] = texto_feriados
        salvar_feriados_locais(st.session_state["feriados_locais_dict"])

    return {
        "banca": banca,
        "local": local,
        "ano": ano,
        "mes_nome": mes_nome,
        "mes_numero": numero_do_mes(mes_nome),
        "dias_permitidos": dias_permitidos,
        "exam_manha": exam_manha,
        "exam_tarde": exam_tarde,
        "pistas": pistas,
        "corte_a": corte_a,
        "itinerante": itinerante,
        "posto_sobra": posto_sobra,
        "id_local": id_local,
    }


def _periodo_padrao_do_gerador(
    p: dict[str, Any], df_completo: pd.DataFrame
) -> tuple[datetime.date, datetime.date]:
    """Mês inteiro; para município itinerante, do primeiro ao último dia de
    viagem no mês."""
    primeiro = datetime.date(p["ano"], p["mes_numero"], 1)
    ultimo = datetime.date(
        p["ano"], p["mes_numero"], calendar.monthrange(p["ano"], p["mes_numero"])[1]
    )
    if p["itinerante"] and not df_completo.empty:
        datas = [d for d in (para_data(x) for x in df_completo["Data"].unique()) if d]
        if datas:
            return min(datas), max(datas)
    return primeiro, ultimo


def _painel_gerador(p: dict[str, Any], chave: str, df_completo: pd.DataFrame) -> None:
    """Gerador automático: demanda do mês, período e dias de cada categoria."""
    primeiro = datetime.date(p["ano"], p["mes_numero"], 1)
    ultimo = datetime.date(
        p["ano"], p["mes_numero"], calendar.monthrange(p["ano"], p["mes_numero"])[1]
    )
    guardadas = opcoes_do_gerador(chave)
    sufixo = f"{_assinatura(chave)[:10]}_{sufixo_widget()}"
    ja_lancado = int(df_completo[COLS_CATEGORIA + [COL_PCD_GERADA]].to_numpy().sum()) if not df_completo.empty else 0

    with st.expander("Gerador automático de calendário", expanded=ja_lancado == 0):
        st.caption(
            "Informe a demanda do mês, o período e os dias de cada categoria. O gerador"
            f" respeita as regras de capacidade: 2 exames por examinador em B a E, Cat A"
            f" em blocos fechados de {VAGAS_CAT_A_POR_DUPLA} por pista (dupla), metade de"
            f" B a E em {', '.join(horarios_fim_turno()) or 'nenhum horário'}, A/C/D/E e"
            " PCD B só de manhã. As demais colunas PCD continuam manuais."
        )

        # Período do mês. Só fica guardado quando foge do padrão (mês inteiro
        # ou dias da viagem): assim, se a viagem mudar de data, o período
        # acompanha em vez de ficar preso à data antiga.
        periodo_padrao = _periodo_padrao_do_gerador(p, df_completo)
        padrao_ini, padrao_fim = (
            _periodo_das_opcoes(guardadas, p["ano"], p["mes_numero"])
            if guardadas.get("periodo")
            else periodo_padrao
        )
        periodo = st.date_input(
            "Período da geração",
            value=(padrao_ini, padrao_fim),
            min_value=primeiro,
            max_value=ultimo,
            format="DD/MM/YYYY",
            key=f"ger_periodo_{sufixo}",
            help="Linhas fora do período não são alteradas: dá para gerar o mês em partes.",
        )
        if not (isinstance(periodo, (tuple, list)) and len(periodo) == 2):
            st.info("Selecione a data inicial e a final do período.")
            return
        periodo = (periodo[0], periodo[1])

        # Demanda por categoria.
        st.markdown("**Demanda do período**")
        chaves = CATEGORIAS + [COL_PCD_GERADA]
        demanda: dict[str, int] = {}
        for coluna, chave_cat in zip(st.columns(len(chaves)), chaves):
            demanda[chave_cat] = int(
                coluna.number_input(
                    chave_cat if chave_cat == COL_PCD_GERADA else f"Cat {chave_cat}",
                    min_value=0,
                    max_value=100000,
                    value=_inteiro(guardadas.get(chave_cat)),
                    step=12 if chave_cat == "A" else 10,
                    key=f"ger_{chave_cat.replace(' ', '_')}_{sufixo}",
                )
            )

        # Dias da semana de cada categoria.
        dias_validos = (
            DIAS_SEMANA_OPCOES
            if p["itinerante"]
            else [d for d in DIAS_SEMANA_OPCOES if d in p["dias_permitidos"]] or DIAS_SEMANA_OPCOES
        )
        dias_guardados = guardadas.get("dias") or {}
        dias_por_categoria: dict[str, list[str]] = {}
        with st.container(border=True):
            st.markdown("**Dias de atendimento de cada categoria**")
            colunas_dias = st.columns(3)
            for indice, chave_cat in enumerate(chaves):
                padrao = dias_guardados.get(chave_cat)
                if not isinstance(padrao, list):
                    padrao = list(DIAS_PCD_PADRAO) if chave_cat == COL_PCD_GERADA else list(dias_validos)
                padrao = [d for d in padrao if d in dias_validos]
                with colunas_dias[indice % 3]:
                    escolhidos = st.pills(
                        chave_cat if chave_cat == COL_PCD_GERADA else f"Cat {chave_cat}",
                        dias_validos,
                        selection_mode="multi",
                        default=padrao,
                        format_func=lambda d: d[:3],
                        key=f"ger_dias_{chave_cat.replace(' ', '_')}_{sufixo}",
                    )
                    dias_por_categoria[chave_cat] = [d for d in dias_validos if d in (escolhidos or [])]

        novas = {
            **demanda,
            "periodo": (
                [periodo[0].isoformat(), periodo[1].isoformat()]
                if periodo != periodo_padrao
                else None
            ),
            "dias": dias_por_categoria,
        }
        # Só abrir a localidade não grava nada: as opções são guardadas quando
        # há demanda ou quando já existiam.
        if novas != guardadas and (guardadas or any(demanda.values())):
            st.session_state["demandas_dict"][chave] = novas
            salvar_demandas(st.session_state["demandas_dict"])

        sem_dia = [
            c for c in chaves if demanda[c] and not dias_por_categoria.get(c)
        ]
        if sem_dia:
            st.warning("Com demanda e sem nenhum dia marcado: " + ", ".join(sem_dia) + ".")

        if p["posto_sobra"]:
            st.info(
                f"**{p['local']}** é posto de sobra: os examinadores de cada dia são os que"
                " sobram do efetivo depois das demais localidades e das viagens"
                + (" (toda a sobra)." if usa_toda_sobra(p["banca"], p["local"]) else
                   f" (até {p['exam_manha']} de manhã e {p['exam_tarde']} à tarde).")
                + " Gere este posto depois de montar as outras localidades."
            )

        no_periodo = df_completo[
            pd.to_datetime(df_completo["Data"], format="%d/%m/%Y").dt.date.between(*periodo)
        ] if not df_completo.empty else df_completo
        lancado_periodo = (
            int(no_periodo[COLS_CATEGORIA + [COL_PCD_GERADA]].to_numpy().sum())
            if not no_periodo.empty
            else 0
        )
        substituir = True
        if lancado_periodo:
            substituir = st.checkbox(
                f"Substituir as {lancado_periodo} vagas (Cat A a E e PCD B) já lançadas no período",
                key=f"ger_substituir_{sufixo}",
            )
        col_gerar, col_sobra = st.columns(2)
        if col_gerar.button(
            "Gerar calendário", key=f"ger_btn_{sufixo}", type="primary", icon=":material/auto_awesome:"
        ):
            if not substituir:
                st.warning("Marque a confirmação para substituir as vagas já lançadas.")
                return
            if not any(demanda.values()):
                st.warning("Informe a demanda de pelo menos uma categoria.")
                return
            exam_por_dia = (
                examinadores_do_posto_de_sobra(p["banca"], p["local"], p["mes_nome"], p["ano"])
                if p["posto_sobra"]
                else None
            )
            resultado = gerar_calendario_do_local(
                p["banca"],
                p["local"],
                p["ano"],
                p["mes_numero"],
                novas,
                exam_manha=p["exam_manha"],
                exam_tarde=p["exam_tarde"],
                pistas=p["pistas"],
                examinadores_por_dia=exam_por_dia,
            )
            if resultado is None:
                st.warning("Nenhuma data de atendimento para gerar.")
                return
            for tipo, mensagem in resumo_da_geracao(p["local"], resultado):
                _flash(tipo, mensagem)
            dados = dados_da_banca_no_mes(p["banca"], p["mes_nome"], p["ano"])
            alertas = alertas_de_sobra(p["banca"], p["mes_nome"], p["ano"], dados)
            if alertas:
                _flash("error", texto_alerta_sobra(p["banca"], alertas))
            ja_gerados = [
                local
                for local in postos_de_sobra(st.session_state["bancas_config"], p["banca"])
                if local != p["local"]
                and local in dados
                and int(dados[local]["Total"].sum()) > 0
            ]
            if ja_gerados and not p["posto_sobra"]:
                _flash(
                    "info",
                    "Os postos de sobra (" + ", ".join(ja_gerados) + ") já estavam gerados"
                    " neste mês. Gere-os de novo (Quadro Geral > Fechamento do mês) para"
                    " que usem a sobra atualizada.",
                )
            st.rerun()
        if p["posto_sobra"] and col_sobra.button(
            "Gerar todos os postos de sobra da banca",
            key=f"ger_sobra_{sufixo}",
            icon=":material/auto_mode:",
            help="Gera os postos de sobra em ordem (número fixo primeiro, depois toda a sobra).",
        ):
            for tipo, mensagem in gerar_postos_de_sobra(p["banca"], p["mes_nome"], p["ano"]):
                _flash(tipo, mensagem)
            st.rerun()


def _calendario_visual(
    banca: str,
    local: str,
    ano: int,
    mes_numero: int,
    feriados: set[datetime.date],
    chave: str,
) -> None:
    if not TEM_COMPONENTE_CALENDARIO:
        st.info(
            "Para arrastar viagens diretamente na tela, instale o pacote"
            " `streamlit-calendar` (já listado no `requirements.txt`)."
        )
        return

    eventos = []
    for viagem in st.session_state["viagens_registradas"]:
        if not banca_vinculada(viagem, banca):
            continue
        inicio = para_data(viagem["Data Inicio"])
        fim = para_data(viagem["Data Fim"])
        if not inicio or not fim:
            continue
        eventos.append(
            {
                "id": str(viagem.get("id") or ""),
                "title": (
                    f"🚍 {viagem['Destino']}"
                    f" ({examinadores_da_banca_na_viagem(viagem, banca)} Ex.)"
                ),
                "start": inicio.isoformat(),
                "end": (fim + datetime.timedelta(days=1)).isoformat(),
                "color": COR_SECUNDARIA if viagem["Destino"] == local else COR_NEUTRA,
                "allDay": True,
                "editable": True,
            }
        )
    for feriado in feriados:
        if feriado.month == mes_numero and feriado.year == ano:
            eventos.append(
                {
                    "id": "",
                    "title": "🔴 Feriado / Sem Atendimento",
                    "start": feriado.isoformat(),
                    "color": COR_ALERTA,
                    "allDay": True,
                    "editable": False,
                }
            )

    opcoes = {
        "editable": True,
        "selectable": False,
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,timeGridWeek",
        },
        "initialDate": f"{ano}-{mes_numero:02d}-01",
        "locale": "pt-br",
    }
    # A versão na chave zera o componente depois de cada arraste tratado; sem
    # isso o mesmo "eventChange" voltava a cada rerun.
    versao = int(st.session_state.get("_ver_cal_visual", 0))
    retorno = componente_calendario(
        events=eventos, options=opcoes, key=f"cal_visual_{_assinatura(chave)[:10]}_{versao}"
    )
    alteracao = (retorno or {}).get("eventChange")
    if not alteracao:
        return
    st.session_state["_ver_cal_visual"] = versao + 1

    evento = (alteracao or {}).get("event") or {}
    id_viagem = evento.get("id")
    if not id_viagem:
        st.rerun()
    try:
        novo_inicio = datetime.date.fromisoformat(str(evento["start"])[:10])
        novo_fim = datetime.date.fromisoformat(str(evento.get("end") or evento["start"])[:10])
        if evento.get("end"):
            novo_fim -= datetime.timedelta(days=1)
    except (KeyError, TypeError, ValueError):
        LOG.warning("Evento de calendário com datas ilegíveis: %r", evento)
        st.rerun()
    novo_fim = max(novo_fim, novo_inicio)

    for viagem in st.session_state["viagens_registradas"]:
        if str(viagem.get("id")) != str(id_viagem):
            continue
        candidata = dict(viagem)
        candidata["Data Inicio"] = novo_inicio
        candidata["Data Fim"] = novo_fim
        candidata["Turnos Por Data"] = {
            iso: turno
            for iso, turno in (viagem.get("Turnos Por Data") or {}).items()
            if (d := para_data(iso)) and novo_inicio <= d <= novo_fim
        }
        erros = validar_viagem(candidata, ignorar_id=viagem.get("id"))
        if erros:
            for erro in erros:
                _flash("error", f"Reagendamento recusado: {erro}")
        else:
            viagem.update(candidata)
            salvar_viagens(st.session_state["viagens_registradas"])
            _flash("success", f"Viagem para {viagem['Destino']} reagendada.")
        break
    st.rerun()


def _monitor_efetivo(
    banca: str,
    datas_do_mes: list[str],
    totais: dict[datetime.date, dict[str, Any]],
) -> None:
    st.sidebar.markdown("---")
    st.sidebar.subheader("Consulta de efetivo por data")
    if not datas_do_mes:
        st.sidebar.info("Nenhuma data de atendimento neste mês.")
        return
    data_consulta = st.sidebar.selectbox(
        "Data para checar o total:", datas_do_mes, key="cal_data_consulta"
    )
    data = para_data(data_consulta)
    if not data:
        return
    efetivo = efetivo_do_dia(banca, data)
    valores = totais.get(data, {"M": 0, "T": 0, "locais": {}, "equipes": {}})
    detalhes = [
        {"Localidade": local, "Manhã": m, "Tarde": t}
        for local, (m, t) in sorted(valores["locais"].items())
    ]
    for (banca_eq, numero), (m, t) in sorted(valores["equipes"].items()):
        detalhes.append(
            {"Localidade": f"Itinerante {numero:02d} ({banca_eq})", "Manhã": m, "Tarde": t}
        )
    pico = pico_diario(valores["M"], valores["T"])
    st.sidebar.markdown(f"**Escala do dia:** `{data_consulta}`")
    st.sidebar.metric("Efetivo usado no dia", f"{pico} / {efetivo}")
    st.sidebar.metric("Saldo restante", efetivo - pico)
    if pico > efetivo:
        st.sidebar.error(f"Efetivo excedido em {pico - efetivo} examinador(es).")
    if detalhes:
        st.sidebar.dataframe(pd.DataFrame(detalhes), **LARGURA_TOTAL, hide_index=True)
    else:
        st.sidebar.info("Sem exames ou viagens nesta data.")


def _renderizar_modo_foco_calendario() -> None:
    """Tela dedicada à montagem do calendário, sem as outras páginas."""
    col_titulo, col_sair = st.columns([5, 1.4])
    with col_titulo:
        cabecalho_pagina(
            "Modo de montagem do calendário",
            "As outras páginas ficam escondidas enquanto você trabalha aqui. Os"
            " lançamentos são salvos a cada edição.",
        )
    with col_sair:
        st.write("")
        if st.button(
            "Concluir e voltar",
            key="btn_sair_modo_foco",
            icon=":material/check:",
            **LARGURA_TOTAL,
        ):
            st.session_state["modo_foco_calendario"] = False
            st.rerun()
    pagina_calendario()


def _faixa_viagens_do_local(banca: str, local: str) -> None:
    viagens_do_local = [
        v
        for v in st.session_state["viagens_registradas"]
        if v.get("Banca") == banca and v.get("Destino") == local
    ]
    for v in sorted(viagens_do_local, key=lambda v: para_data(v["Data Inicio"])):
        ajustes = descricao_turnos(v)
        chegada, retorno = horarios_limite_da_viagem(v)
        st.success(
            f"Viagem {int(v.get('Numero Banca Itinerante', 0) or 0):02d}: {local} de"
            f" **{formatar_br(para_data(v['Data Inicio']))}** a"
            f" **{formatar_br(para_data(v['Data Fim']))}** com"
            f" **{total_examinadores(v)} examinador(es)** ({v.get('Turno', TURNO_INTEGRAL).lower()})"
            + (f" · provas a partir de **{chegada}** no 1º dia" if chegada else "")
            + (f" · até **{retorno}** no último dia" if retorno else "")
            + (f" · ajustes: {ajustes}" if ajustes else "")
            + ".",
            icon=":material/directions_bus:",
        )


def pagina_calendario() -> None:
    p = _barra_lateral_calendario()
    if p is None:
        return
    banca, local, ano = p["banca"], p["local"], p["ano"]
    mes_nome, mes_numero = p["mes_nome"], p["mes_numero"]

    if not st.session_state.get("modo_foco_calendario"):
        col_titulo, col_botao = st.columns([4, 1.4])
        with col_titulo:
            cabecalho_pagina(
                f"Calendário · {local}",
                f"{banca} · {mes_nome}/{ano}"
                + (" · município itinerante" if p["itinerante"] else "")
                + (" · posto de sobra" if p["posto_sobra"] else ""),
            )
        col_botao.write("")
        if col_botao.button(
            "Tela cheia",
            key="btn_entrar_modo_foco",
            icon=":material/fullscreen:",
            help="Monta o calendário sem as outras páginas por perto.",
            **LARGURA_TOTAL,
        ):
            st.session_state["modo_foco_calendario"] = True
            st.rerun()

    if not horarios_ativos():
        st.warning("Não há horários ativos. Cadastre ou reative horários em **Horários & Regras**.")
        return
    if not p["itinerante"] and not p["dias_permitidos"]:
        st.warning("Selecione ao menos um dia da semana na barra lateral.")
        return

    chave = chave_historico(banca, mes_nome, ano, local)
    df_base, df_completo, datas_do_mes = grade_do_local(
        banca, local, ano, mes_numero, padrao_manha=p["exam_manha"], padrao_tarde=p["exam_tarde"]
    )
    feriados = feriados_do_local(banca, local, ano)

    if p["itinerante"]:
        _faixa_viagens_do_local(banca, local)
        if df_completo.empty:
            st.info(
                f"**{local}** é atendido por equipe itinerante e não há viagem de"
                f" {banca} para cá em {mes_nome}/{ano}. Cadastre a viagem em **Viagens"
                " Itinerantes**: a grade sai com as datas, os turnos e os horários dela."
            )
            return
    elif df_completo.empty:
        st.warning("Nenhuma data de atendimento gerada para os parâmetros atuais.")
        return

    _painel_gerador(p, chave, df_completo)

    # O componente de calendário devolve um evento ao montar, o que roda a
    # página de novo e apagava os avisos da geração: só é montado sob pedido.
    if st.toggle(
        "Mostrar calendário visual (arraste para reagendar viagens)",
        key=f"cal_visual_ligado_{p['id_local']}",
    ):
        _calendario_visual(banca, local, ano, mes_numero, feriados, chave)

    st.markdown('<div class="titulo-secao">Lançamento e edição de vagas por horário</div>', unsafe_allow_html=True)
    total_dias = calendar.monthrange(ano, mes_numero)[1]
    primeiro_dia = datetime.date(ano, mes_numero, 1)
    ultimo_dia = datetime.date(ano, mes_numero, total_dias)
    col_periodo, col_datas = st.columns([1, 2])
    with col_periodo:
        periodo = st.date_input(
            "Filtrar por período:",
            value=(primeiro_dia, ultimo_dia),
            min_value=primeiro_dia,
            max_value=ultimo_dia,
            format="DD/MM/YYYY",
            key=f"cal_filtro_periodo_{p['id_local']}_{mes_numero}_{ano}",
        )
    with col_datas:
        datas_selecionadas = st.multiselect(
            "Ou selecione dias específicos do mês:",
            options=datas_do_mes,
            default=[],
            placeholder="Selecione um ou mais dias...",
            key=f"cal_filtro_datas_{p['id_local']}_{mes_numero}_{ano}",
        )

    if datas_selecionadas:
        df_exibicao = df_completo[df_completo["Data"].isin(datas_selecionadas)].copy()
    elif isinstance(periodo, (tuple, list)) and len(periodo) == 2:
        inicio, fim = periodo
        convertidas = pd.to_datetime(df_completo["Data"], format="%d/%m/%Y").dt.date
        df_exibicao = df_completo[(convertidas >= inicio) & (convertidas <= fim)].copy()
    else:
        df_exibicao = df_completo.copy()

    # A chave do editor inclui o desenho da grade (datas x horários): se o
    # desenho muda (dia, horário, feriado, viagem, filtro), as edições
    # guardadas por posição são descartadas em vez de caírem na linha errada.
    desenho = assinatura_dados(df_exibicao[["Data", "Horário"]].values.tolist())[:10]
    chave_editor = (
        f"editor_{_assinatura(chave)[:10]}_{desenho}_{_versao_editor(chave)}_{sufixo_widget()}"
    )
    df_previo = _normalizar_grade(_aplicar_edicoes_pendentes(df_exibicao, chave_editor))
    df_previo = colunas_de_validacao(df_previo, p["pistas"], p["corte_a"])

    configuracao = {
        "Data": st.column_config.TextColumn("Data", disabled=True),
        "Dia da Semana": st.column_config.TextColumn("Dia", disabled=True),
        "Status": st.column_config.SelectboxColumn("Status", options=OPCOES_STATUS, required=True),
        "Exam. M": st.column_config.NumberColumn("Exam. M", min_value=0, max_value=100, step=1, default=0),
        "Exam. T": st.column_config.NumberColumn("Exam. T", min_value=0, max_value=100, step=1, default=0),
        "Horário": st.column_config.TextColumn("Horário", disabled=True),
        "Total": st.column_config.NumberColumn("Total", disabled=True),
        "Exam. Necess.": st.column_config.NumberColumn(
            "Exam. necess.", disabled=True, format="%.1f",
            help="Examinadores que as vagas do horário exigem.",
        ),
        "Situação": st.column_config.TextColumn("Situação", disabled=True, width="medium"),
        "Horário Quebrado": st.column_config.TextColumn(
            "Horário quebrado", disabled=True,
            help="C, D e E dividindo o horário com a Cat B saem +1, +2 e +10 min.",
        ),
    }
    for coluna in COLS_VAGAS:
        configuracao[coluna] = st.column_config.NumberColumn(coluna, min_value=0, step=1, default=0)
    configuracao["Cat A"] = st.column_config.NumberColumn(
        "Cat A", min_value=0, step=6, default=0,
        help=f"Blocos de {VAGAS_CAT_A_POR_DUPLA} por pista (dupla); 6 no fim de turno com corte.",
    )
    # Situação e examinadores necessários ficam logo no começo, à vista,
    # ao lado dos examinadores escalados.
    ordem = (
        ["Data", "Dia da Semana", "Horário", "Situação", "Status", "Exam. M", "Exam. T", "Exam. Necess."]
        + COLS_VAGAS
        + ["Total", "Horário Quebrado"]
    )

    df_editado = st.data_editor(
        _estilizar_grade_calendario(df_previo),
        column_config=configuracao,
        column_order=ordem,
        **LARGURA_TOTAL,
        num_rows="fixed",
        height=450,
        key=chave_editor,
    )
    st.caption(
        "Verde: com vagas · cinza: bloqueada · vermelho: capacidade estourada ou regra"
        " violada · sem cor: vazia. Linhas bloqueadas pelo sistema (feriado, horário"
        " fora da viagem, dia fora da grade) só são liberadas removendo a causa; as"
        " vagas lançadas nelas ficam guardadas e voltam quando o bloqueio sai."
    )

    df_editado = _normalizar_grade(df_editado[[c for c in COLUNAS_GRADE if c in df_editado.columns]])
    df_completo = df_completo.copy()
    df_completo.update(df_editado[COLUNAS_EDITAVEIS])
    # Status imposto pelo sistema não pode ser trocado na grade.
    automaticas = df_completo[COL_AUTO] == 1
    status_base = df_base.drop(columns=[COL_VIAGEM]).set_index(["Data", "Horário"])["Status"]
    for indice in df_completo.index[automaticas]:
        chave_linha = (df_completo.at[indice, "Data"], df_completo.at[indice, "Horário"])
        if chave_linha in status_base.index:
            df_completo.at[indice, "Status"] = status_base.loc[chave_linha]
    df_completo = _normalizar_grade(df_completo)
    _guardar_grade(chave, df_completo)

    problemas = violacoes(df_completo, p["pistas"], p["corte_a"])
    if problemas:
        st.error(
            f"{len(problemas)} horário(s) com problema de capacidade:\n\n"
            + "\n".join(f"- {m}" for m in problemas[:15])
            + ("\n- ..." if len(problemas) > 15 else "")
        )

    dados_banca = dados_da_banca_no_mes(banca, mes_nome, ano)
    totais = efetivo_diario_da_banca(banca, mes_nome, ano, dados_banca)
    efetivos = efetivo_do_mes(banca, ano, mes_numero)
    _monitor_efetivo(banca, datas_do_mes, totais)

    picos = {d: pico_diario(v["M"], v["T"]) for d, v in totais.items()}
    estouros = sorted(d for d, pico in picos.items() if pico > efetivos.get(d, 0))
    folgas = [efetivos[d] - picos.get(d, 0) for d in efetivos if d.weekday() < 5]
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Total de vagas da localidade", f"{int(df_completo['Total'].sum())} vagas")
    col_b.metric("Maior pico diário da banca", f"{max(picos.values(), default=0)} exam.")
    col_c.metric("Menor folga diária", f"{min(folgas, default=0)} exam.")

    alertas = alertas_de_sobra(banca, mes_nome, ano, dados_banca)
    if alertas:
        st.error(texto_alerta_sobra(banca, alertas), icon=":material/warning:")

    problemas_banca = violacoes_da_banca(banca, dados_banca)
    if estouros or problemas_banca:
        motivos = []
        if estouros:
            motivos.append(
                "o efetivo do dia é excedido em "
                + ", ".join(f"{formatar_curto(d)} ({picos[d]}/{efetivos[d]})" for d in estouros[:10])
            )
        if problemas_banca:
            motivos.append(f"há {len(problemas_banca)} horário(s) com capacidade estourada na banca")
        st.error(
            "**Bloqueio de segurança**: " + " e ".join(motivos)
            + ". Ajuste a escala antes de gerar os documentos.",
            icon=":material/block:",
        )
        return

    st.markdown('<div class="titulo-secao">Conferência e impressão</div>', unsafe_allow_html=True)
    col_pdf_local, col_pdf_banca = st.columns(2)
    with col_pdf_local:
        st.markdown(f"**Grade detalhada: {local}**")
        bloco_pdf(
            chave=f"loc_{chave}",
            rotulo_gerar="Gerar PDF desta localidade",
            gerador=lambda: gerar_pdf_localidade(df_completo, banca, local, mes_nome, ano),
            nome_arquivo=f"Calendario_{banca}_{local}_{mes_nome}_{ano}.pdf".replace(" ", "_"),
            assinatura=assinatura_dados(df_completo),
            ajuda="Linha a linha, por horário, apenas desta localidade.",
        )
    with col_pdf_banca:
        st.markdown(f"**Calendário completo: {banca}**")
        st.caption(
            f"{len(dados_banca)} localidade(s) lançada(s) no mês. Ordem: sede,"
            " região metropolitana e demais municípios por data do primeiro exame."
        )
        bloco_pdf(
            chave=f"banca_{banca}_{mes_nome}_{ano}",
            rotulo_gerar="Gerar PDF completo da banca",
            gerador=lambda: gerar_pdf_banca(
                dados_banca,
                st.session_state["viagens_registradas"],
                st.session_state["bancas_config"],
                banca,
                mes_nome,
                ano,
                mes_numero,
                efetivos,
            ),
            nome_arquivo=f"Calendario_Completo_{banca}_{mes_nome}_{ano}.pdf".replace(" ", "_"),
            assinatura=assinatura_dados(
                dados_banca,
                [_sem_id(v) for v in st.session_state["viagens_registradas"]],
                {d.isoformat(): n for d, n in efetivos.items()},
            ),
            ajuda="Calendário da banca inteira, com a conferência de efetivo diário.",
        )


# --- Página: gestão de viagens itinerantes -----------------------------------
def _validar_capacidade(candidata: Viagem, ignorar_id: Any = None) -> list[str]:
    """Efetivo de cada banca envolvida em todos os dias do trecho.

    Conta todos os dias entre a ida e a volta (o deslocamento ocupa a equipe
    o dia inteiro), identifica equipes por (banca titular, número) e checa:
      * deslocamento simultâneo x teto fora da sede e efetivo do dia;
      * sobra mínima para os postos montados por último (Pátio DETRAN e
        Castelinho; Imperatriz), somando as viagens às demais localidades.
    """
    erros: list[str] = []
    limites = st.session_state["limite_fora_sede_bancas"]
    capacidades = st.session_state["capacidade_bancas"]
    outras = [
        v
        for v in st.session_state["viagens_registradas"]
        if ignorar_id is None or v.get("id") != ignorar_id
    ]
    cenario = outras + [candidata]
    inicio = para_data(candidata.get("Data Inicio"))
    fim = para_data(candidata.get("Data Fim"))
    if not inicio or not fim:
        return erros

    for banca in bancas_da_viagem(candidata):
        if examinadores_da_banca_na_viagem(candidata, banca) <= 0:
            continue
        for dia in dias_do_intervalo(inicio, fim):
            manha, tarde = efetivo_em_viagem(cenario, banca, dia)
            pico = pico_diario(manha, tarde)
            limite = int(limites.get(banca, capacidades.get(banca, 0)) or 0)
            teto = min(limite, efetivo_do_dia(banca, dia))
            if pico > teto:
                erros.append(
                    f"{banca}: {pico} examinadores fora da sede em {formatar_br(dia)}"
                    f" excede o teto de {teto} (máximo fora da sede ou efetivo do dia)."
                )
                break

        meses = sorted({(d.year, d.month) for d in dias_do_intervalo(inicio, fim)})
        for ano, mes_numero in meses:
            alertas = [
                (d, n)
                for d, n in alertas_de_sobra(banca, MESES_LISTA[mes_numero - 1], ano, viagens=cenario)
                if inicio <= d <= fim
            ]
            if alertas:
                erros.append(f"{banca}: " + texto_alerta_sobra(banca, alertas))
                break
    return erros


def validar_viagem(candidata: Viagem, ignorar_id: Any = None) -> list[str]:
    """Regras de uma viagem (ou trecho), usadas no cadastro, na edição e no
    arraste do calendário. Antes só o cadastro validava."""
    erros: list[str] = []
    inicio = para_data(candidata.get("Data Inicio"))
    fim = para_data(candidata.get("Data Fim"))
    if not inicio or not fim:
        return ["Informe as datas de início e término."]
    if fim < inicio:
        return ["A data final não pode ser anterior à data de início."]
    if total_examinadores(candidata) <= 0:
        erros.append("Informe ao menos um examinador para a viagem.")
    dias = dias_efetivos_da_viagem(candidata)
    if not dias:
        erros.append(
            f"Todos os dias do período estão como '{TURNO_SEM_ATENDIMENTO}'."
            " A viagem não teria atendimento."
        )
    chegada, retorno = horarios_limite_da_viagem(candidata)
    if dias and len(dias) == 1 and chegada and retorno:
        if (_hora_em_minutos(chegada) or 0) > (_hora_em_minutos(retorno) or 0):
            erros.append("Com um só dia de atendimento, a chegada não pode ser depois do retorno.")
    if dias and not any(
        viagem_cobre_horario(candidata, d, h) for d in dias for h in horarios_ativos()
    ):
        erros.append(
            "Nenhum horário de exame fica dentro do período, dos turnos e dos horários"
            " de chegada e retorno informados."
        )
    conflitos = conflitos_de_rota(
        st.session_state["viagens_registradas"],
        banca=candidata.get("Banca", ""),
        numero_equipe=int(candidata.get("Numero Banca Itinerante", 0) or 0),
        inicio=inicio,
        fim=fim,
        turno=candidata.get("Turno", TURNO_INTEGRAL),
        turnos_por_data=candidata.get("Turnos Por Data") or {},
        ignorar_id=ignorar_id,
    )
    if conflitos:
        destinos = ", ".join(sorted({str(c.get("Destino", "?")) for c in conflitos}))
        erros.append(
            f"Esta equipe já atende {destinos} no mesmo turno de algum dia deste"
            " período. Use o ajuste de turno por dia para separar manhã e tarde."
        )
    if st.session_state.get("controle_capacidade_ativo"):
        erros.extend(_validar_capacidade(candidata, ignorar_id))
    return erros


def _alertas_de_feriado(
    banca: str, destino: str, inicio: datetime.date, fim: datetime.date
) -> list[datetime.date]:
    """Feriados (nacionais, estaduais, facultativos marcados e municipais)
    dentro do período da viagem."""
    datas: set[datetime.date] = set()
    for ano in range(inicio.year, fim.year + 1):
        datas |= feriados_do_local(banca, destino, ano)
    return sorted(d for d in datas if inicio <= d <= fim)


def _editor_turnos_por_dia(
    inicio: datetime.date,
    fim: datetime.date,
    turno_geral: str,
    prefixo_chave: str,
    valores_atuais: dict[str, str] | None = None,
    expandido: bool = False,
) -> dict[str, str]:
    """Turno dia a dia dentro do período da viagem (ex.: Pinheiro de manhã e
    São Bento à tarde no dia 12, sem contar a equipe duas vezes)."""
    valores_atuais = valores_atuais or {}
    dias = dias_do_intervalo(inicio, fim)
    if not dias:
        return {}
    if len(dias) > 45:
        st.warning("Período muito longo para o ajuste dia a dia.")
        return dict(valores_atuais)

    # Se o período ou o turno geral mudam, os seletores antigos são
    # descartados para não acusarem conflito onde não há.
    chave_contexto = f"{prefixo_chave}__contexto"
    contexto = (inicio.isoformat(), fim.isoformat(), turno_geral)
    if st.session_state.get(chave_contexto) != contexto:
        for chave_antiga in [
            k for k in st.session_state if k.startswith(f"{prefixo_chave}_") and k != chave_contexto
        ]:
            st.session_state.pop(chave_antiga, None)
        st.session_state[chave_contexto] = contexto

    resultado: dict[str, str] = {}
    with st.expander(f"Ajuste de turno por dia ({len(dias)} dia(s))", expanded=expandido):
        st.caption(
            "Por padrão todos os dias seguem o período geral. Altere só os dias"
            f" em que a equipe atende um turno, ou use **{TURNO_SEM_ATENDIMENTO}**"
            " para dias só de deslocamento e fins de semana. O deslocamento conta"
            " no efetivo do dia mesmo sem exame."
        )
        colunas = st.columns(min(len(dias), 3))
        for indice, dia in enumerate(dias):
            iso = dia.isoformat()
            atual = valores_atuais.get(iso, turno_geral)
            if atual not in TURNOS_POR_DIA:
                atual = turno_geral
            with colunas[indice % len(colunas)]:
                escolhido = st.selectbox(
                    f"{formatar_curto(dia)} ({nome_dia_semana(dia)[:3]})",
                    TURNOS_POR_DIA,
                    index=TURNOS_POR_DIA.index(atual),
                    key=f"{prefixo_chave}_{iso}",
                )
            if escolhido != turno_geral:
                resultado[iso] = escolhido
    return resultado


SEM_LIMITE_HORARIO = "Sem restrição"


def _seletores_chegada_retorno(
    prefixo: str, atual_chegada: str = "", atual_retorno: str = ""
) -> tuple[str, str]:
    """Primeiro horário de exame no dia de chegada e último no dia de retorno."""
    ativos = horarios_ativos()
    opcoes = [SEM_LIMITE_HORARIO] + ativos
    col_chegada, col_retorno = st.columns(2)
    chegada = col_chegada.selectbox(
        "1ª turma no dia da ida:",
        opcoes,
        index=opcoes.index(atual_chegada) if atual_chegada in opcoes else 0,
        key=f"{prefixo}_chegada",
        help=(
            "Horário da primeira turma no dia de chegada. Ex.: Codó fica a 5 h de"
            " viagem; saindo de manhã, a primeira turma é às 13:30."
        ),
    )
    retorno = col_retorno.selectbox(
        "Última turma na volta:",
        opcoes,
        index=opcoes.index(atual_retorno) if atual_retorno in opcoes else 0,
        key=f"{prefixo}_retorno",
        help="Horário da última turma no dia da volta, para dar tempo do deslocamento.",
    )
    return (
        "" if chegada == SEM_LIMITE_HORARIO else chegada,
        "" if retorno == SEM_LIMITE_HORARIO else retorno,
    )


def _formulario_cadastro_viagem() -> None:
    st.markdown('<div class="titulo-secao">Cadastrar nova viagem</div>', unsafe_allow_html=True)
    bancas_config = st.session_state["bancas_config"]
    if not bancas_config:
        st.warning("Cadastre ao menos uma banca na página de Gestão de Localidades.")
        return

    viagens = st.session_state["viagens_registradas"]
    banca = st.selectbox("Banca responsável:", list(bancas_config.keys()), key="v_banca")

    proximo = proximo_numero_equipe(viagens, banca)
    opcoes_equipe = {f"Nova banca itinerante (Banca {proximo:02d})": proximo}
    for numero in numeros_de_equipe(viagens, banca):
        opcoes_equipe[f"Continuar Banca {numero:02d} já cadastrada"] = numero
    escolha = st.selectbox(
        "Esta viagem pertence a:",
        list(opcoes_equipe.keys()),
        key="v_equipe",
        help=(
            "Se a mesma equipe vai atender mais de um município na mesma viagem,"
            " cadastre o primeiro trecho como 'Nova banca itinerante' e os"
            " seguintes como 'Continuar'. Assim os examinadores contam uma vez só."
        ),
    )
    numero_equipe = opcoes_equipe[escolha]

    trecho_referencia = next(
        (v for v in viagens if chave_equipe(v) == (banca, numero_equipe)), None
    )
    padrao_examinadores = max(
        1, examinadores_da_banca_na_viagem(trecho_referencia, banca) if trecho_referencia else 2
    )

    bancas_apoio: list[str] = []
    possiveis = bancas_apoio_disponiveis(banca)
    if possiveis:
        bancas_apoio = st.multiselect(
            "Bancas em apoio conjunto:",
            possiveis,
            key="v_apoio",
            help="O efetivo de cada banca é informado separadamente abaixo.",
        )

    destinos = localidades_itinerantes(bancas_config, banca)
    if not destinos:
        st.warning(
            "Esta banca não possui municípios itinerantes cadastrados (os postos da"
            " região metropolitana e a sede não recebem viagem)."
        )
        return
    destino = st.selectbox("Município de destino:", destinos, key="v_destino")

    hoje = datetime.date.today()
    col_ini, col_fim = st.columns(2)
    data_inicio = col_ini.date_input("Ida (1º dia):", hoje, format="DD/MM/YYYY", key="v_inicio")
    data_fim = col_fim.date_input(
        "Volta (último dia):", hoje + datetime.timedelta(days=4), format="DD/MM/YYYY", key="v_fim"
    )
    turno = st.selectbox(
        "Período de atendimento (padrão do trecho):",
        TURNOS,
        key="v_turno",
        help="Padrão para todos os dias. Para dias de um turno só, use o ajuste dia a dia.",
    )
    chegada, retorno = _seletores_chegada_retorno("v_horario")

    turnos_por_data: dict[str, str] = {}
    if data_fim >= data_inicio:
        turnos_por_data = _editor_turnos_por_dia(data_inicio, data_fim, turno, "v_turno_dia")
        if turnos_por_data:
            st.info(
                "Ajustes aplicados: "
                + " · ".join(
                    f"{formatar_curto(para_data(iso))} → {valor}"
                    for iso, valor in sorted(turnos_por_data.items())
                )
            )

    examinadores_por_banca: dict[str, int] = {}
    participantes = bancas_para_validar(banca, bancas_apoio)
    colunas = st.columns(len(participantes))
    for coluna, participante in zip(colunas, participantes):
        examinadores_por_banca[participante] = int(
            coluna.number_input(
                f"Examinadores: {participante.replace('Banca ', '')}",
                min_value=0,
                max_value=100,
                value=min(padrao_examinadores, 100) if participante == banca else 0,
                step=1,
                key=f"v_exam_{participante}",
            )
        )

    observacoes = st.text_input(
        "Observações / Portaria:", placeholder="Ex: Portaria nº 123/2026", key="v_obs"
    )

    controle_ativo = st.checkbox(
        "Controle automático de capacidade (efetivo, máximo fora da sede e sobra mínima)",
        value=bool(st.session_state.get("controle_capacidade_ativo")),
        key=f"chk_controle_{sufixo_widget()}",
        help="Desativado: permite cadastrar viagens sem bloqueio por capacidade.",
    )
    if controle_ativo != bool(st.session_state.get("controle_capacidade_ativo")):
        st.session_state["controle_capacidade_ativo"] = controle_ativo
        salvar_config("controle_capacidade_ativo", controle_ativo)
    st.caption("O efetivo de cada banca e o máximo fora da sede são ajustados na página **Efetivo**.")

    if st.button("Registrar viagem", key="v_btn_salvar", type="primary", icon=":material/save:"):
        candidata = {
            "id": None,
            "Banca": banca,
            "Numero Banca Itinerante": int(numero_equipe),
            "Bancas Apoio": [b for b in bancas_apoio if examinadores_por_banca.get(b)],
            "Destino": destino,
            "Data Inicio": data_inicio,
            "Data Fim": data_fim,
            "Turno": turno,
            "Turnos Por Data": dict(turnos_por_data),
            "Examinadores": sum(examinadores_por_banca.values()),
            "Examinadores Por Banca": {k: int(v) for k, v in examinadores_por_banca.items() if v},
            "Observações": observacoes or "",
            "Horario Inicio": chegada,
            "Horario Fim": retorno,
        }
        erros = validar_viagem(candidata)
        if erros:
            for erro in erros:
                st.error(erro)
            return
        st.session_state["viagens_registradas"].append(candidata)
        salvar_viagens(st.session_state["viagens_registradas"])
        feriados = _alertas_de_feriado(banca, destino, data_inicio, data_fim)
        if feriados:
            _flash(
                "warning",
                "A viagem atravessa feriado(s): "
                + ", ".join(formatar_br(d) for d in feriados)
                + ". Nesses dias a grade da localidade fica bloqueada.",
            )
        _flash("success", f"Trecho para **{destino}** registrado na Banca {int(numero_equipe):02d}.")
        st.rerun()


def _editor_de_trecho(trecho: Viagem) -> None:
    """Edição de um trecho: datas, turnos, horários, efetivo e observação.

    Antes, a edição era por equipe e copiava as datas do primeiro trecho
    para todos os outros. Agora cada trecho tem suas próprias datas e toda
    alteração passa pela mesma validação do cadastro.
    """
    id_trecho = trecho.get("id")
    marca = f"{id_trecho}_{_assinatura(_sem_id(trecho))[:8]}_{sufixo_widget()}"
    col_ini, col_fim, col_turno = st.columns(3)
    inicio = col_ini.date_input(
        "Ida", value=para_data(trecho["Data Inicio"]), format="DD/MM/YYYY", key=f"tr_ini_{marca}"
    )
    fim = col_fim.date_input(
        "Volta", value=para_data(trecho["Data Fim"]), format="DD/MM/YYYY", key=f"tr_fim_{marca}"
    )
    turno_atual = trecho.get("Turno", TURNO_INTEGRAL)
    turno = col_turno.selectbox(
        "Período padrão",
        TURNOS,
        index=TURNOS.index(turno_atual) if turno_atual in TURNOS else 0,
        key=f"tr_turno_{marca}",
    )
    chegada, retorno = _seletores_chegada_retorno(
        f"tr_horario_{marca}", *horarios_limite_da_viagem(trecho)
    )
    ajustes = trecho.get("Turnos Por Data") or {}
    if fim >= inicio:
        ajustes = _editor_turnos_por_dia(inicio, fim, turno, f"tr_dia_{marca}", ajustes)

    por_banca: dict[str, int] = {}
    bancas = bancas_da_viagem(trecho)
    colunas = st.columns(len(bancas))
    for coluna, banca in zip(colunas, bancas):
        por_banca[banca] = int(
            coluna.number_input(
                f"Examinadores: {banca.replace('Banca ', '')}",
                min_value=0,
                max_value=100,
                value=examinadores_da_banca_na_viagem(trecho, banca),
                key=f"tr_exam_{banca}_{marca}",
            )
        )
    observacao = st.text_input(
        "Observações / Portaria", value=str(trecho.get("Observações", "")), key=f"tr_obs_{marca}"
    )

    col_salvar, col_remover = st.columns(2)
    if col_salvar.button("Salvar trecho", key=f"tr_salvar_{marca}", icon=":material/save:"):
        candidata = dict(trecho)
        candidata.update(
            {
                "Data Inicio": inicio,
                "Data Fim": fim,
                "Turno": turno,
                "Turnos Por Data": {
                    iso: v for iso, v in ajustes.items()
                    if (d := para_data(iso)) and inicio <= d <= fim
                },
                "Examinadores Por Banca": {b: n for b, n in por_banca.items() if n},
                "Bancas Apoio": [
                    b for b in (trecho.get("Bancas Apoio") or []) if por_banca.get(b)
                ],
                "Examinadores": sum(por_banca.values()),
                "Observações": observacao,
                "Horario Inicio": chegada,
                "Horario Fim": retorno,
            }
        )
        erros = validar_viagem(candidata, ignorar_id=id_trecho)
        if erros:
            for erro in erros:
                st.error(erro)
        else:
            trecho.update(candidata)
            salvar_viagens(st.session_state["viagens_registradas"])
            _flash("success", f"Trecho para {trecho.get('Destino')} atualizado.")
            st.rerun()
    if col_remover.button("Remover este trecho", key=f"tr_remover_{marca}", icon=":material/delete:"):
        salvar_snapshot(motivo="antes_remover_trecho")
        st.session_state["viagens_registradas"] = [
            v for v in st.session_state["viagens_registradas"] if v is not trecho
        ]
        salvar_viagens(st.session_state["viagens_registradas"])
        _flash("warning", f"Trecho para {trecho.get('Destino')} removido.")
        st.rerun()


def _bloco_equipe(banca: str, numero: int, trechos: list[Viagem]) -> None:
    destinos: list[str] = []
    for trecho in trechos:
        destino = trecho.get("Destino", "")
        if destino and destino not in destinos:
            destinos.append(destino)
    inicio = min(para_data(t["Data Inicio"]) for t in trechos)
    fim = max(para_data(t["Data Fim"]) for t in trechos)
    examinadores = max(examinadores_da_banca_na_viagem(t, banca) for t in trechos)
    with st.container(border=True):
        st.markdown(
            f"**Banca {numero:02d} · {' e '.join(destinos) or 'Destino não informado'}**  \n"
            f"{formatar_br(inicio)} a {formatar_br(fim)} · {examinadores} examinador(es)"
            f" de {banca.replace('Banca ', '')}"
        )

        for trecho in trechos:
            ajustes = descricao_turnos(trecho)
            chegada, retorno = horarios_limite_da_viagem(trecho)
            rotulo = (
                f"{trecho.get('Destino')}: {formatar_curto(para_data(trecho['Data Inicio']))}"
                f" a {formatar_curto(para_data(trecho['Data Fim']))}"
                f" · {trecho.get('Turno', TURNO_INTEGRAL)}"
                + (f" · a partir de {chegada}" if chegada else "")
                + (f" · até {retorno}" if retorno else "")
                + (f" · ajustes: {ajustes}" if ajustes else "")
            )
            with st.expander(rotulo, expanded=False, icon=":material/edit:"):
                _editor_de_trecho(trecho)

        sufixo = f"{banca}_{numero}_{sufixo_widget()}"
        col_num, col_btn, col_remover = st.columns([1.2, 1.4, 1.6])
        novo_numero = int(
            col_num.number_input("Nº da banca", value=int(numero), min_value=1, step=1, key=f"grp_num_{sufixo}")
        )
        col_btn.write("")
        if col_btn.button("Renumerar", key=f"grp_renum_{sufixo}", disabled=novo_numero == numero):
            em_uso = numeros_de_equipe(st.session_state["viagens_registradas"], banca)
            if novo_numero in em_uso:
                st.error(
                    f"A Banca {novo_numero:02d} já existe em {banca}. Escolha um número livre"
                    " (juntar equipes por renumeração foi bloqueado para não somar"
                    " examinadores por engano)."
                )
            else:
                for trecho in trechos:
                    trecho["Numero Banca Itinerante"] = novo_numero
                salvar_viagens(st.session_state["viagens_registradas"])
                _flash("success", f"Banca {numero:02d} renumerada para {novo_numero:02d}.")
                st.rerun()
        col_remover.write("")
        if col_remover.button(f"Remover Banca {numero:02d} inteira", key=f"del_{sufixo}", icon=":material/delete:"):
            salvar_snapshot(motivo="antes_remover_viagem")
            ids = {id(t) for t in trechos}
            st.session_state["viagens_registradas"] = [
                v for v in st.session_state["viagens_registradas"] if id(v) not in ids
            ]
            salvar_viagens(st.session_state["viagens_registradas"])
            _flash("warning", f"Banca {numero:02d} removida.")
            st.rerun()


def _cronograma_viagens() -> None:
    st.markdown('<div class="titulo-secao">Cronograma de viagens programadas</div>', unsafe_allow_html=True)
    viagens = st.session_state["viagens_registradas"]
    if not viagens:
        st.info("Nenhuma viagem cadastrada até o momento.")
        return

    mes_exibicao, ano_exibicao, mes_numero = _seletor_mes_ano("v_cronograma")
    primeiro = datetime.date(ano_exibicao, mes_numero, 1)
    ultimo = datetime.date(ano_exibicao, mes_numero, calendar.monthrange(ano_exibicao, mes_numero)[1])

    st.caption(
        "Escala mensal: o PDF atrela automaticamente 01 preposto e 01 veículo a cada"
        " banca itinerante; essa informação existe apenas no documento."
    )
    bloco_pdf(
        chave=f"escala_viagens_{mes_exibicao}_{ano_exibicao}",
        rotulo_gerar="Gerar PDF da escala de viagens do mês",
        gerador=lambda: gerar_pdf_escala_viagens(
            st.session_state["viagens_registradas"], mes_exibicao, ano_exibicao, mes_numero
        ),
        nome_arquivo=f"Escala_Viagens_{mes_exibicao}_{ano_exibicao}.pdf",
        assinatura=assinatura_dados([_sem_id(v) for v in viagens], mes_exibicao, ano_exibicao),
        ajuda="Consolida todas as equipes itinerantes do mês.",
        altura_previa=560,
    )

    # Só as equipes com algum dia de atendimento no mês escolhido (antes a
    # lista mostrava todas as viagens já cadastradas, de qualquer mês).
    grupos = [
        g
        for g in agrupar_trechos_por_equipe(viagens)
        if any(
            primeiro <= d <= ultimo for t in g["Trechos"] for d in dias_do_intervalo(
                para_data(t["Data Inicio"]), para_data(t["Data Fim"])
            )
        )
    ]
    if not grupos:
        st.info(f"Nenhuma viagem em {mes_exibicao}/{ano_exibicao}.")
        return
    for banca in sorted({g["Banca"] for g in grupos}, key=lambda b: chave_ordem_publicacao(b, "", {})):
        st.markdown(
            f'<div class="faixa-banca">{html.escape(banca)} · {html.escape(mes_exibicao)}'
            f" de {ano_exibicao}</div>",
            unsafe_allow_html=True,
        )
        for grupo in sorted((g for g in grupos if g["Banca"] == banca), key=lambda g: g["Numero"]):
            _bloco_equipe(banca, grupo["Numero"], grupo["Trechos"])


def pagina_viagens() -> None:
    cabecalho_pagina(
        "Viagens Itinerantes",
        "Deslocamentos das bancas. A equipe conta no efetivo todos os dias da ida à"
        " volta, inclusive o deslocamento, e uma única vez por dia.",
    )
    coluna_form, coluna_lista = st.columns([1, 1.6], gap="large")
    with coluna_form:
        _formulario_cadastro_viagem()
    with coluna_lista:
        _cronograma_viagens()


# --- Página: efetivo de examinadores ----------------------------------------
def pagina_efetivo() -> None:
    cabecalho_pagina(
        "Efetivo de Examinadores",
        "Quantos examinadores cada banca tem por dia. É a base do quadro, do"
        " calendário, das viagens e da sobra dos postos montados por último.",
    )
    bancas_config = st.session_state["bancas_config"]
    if not bancas_config:
        st.info("Cadastre uma banca na página Localidades.")
        return
    sufixo = sufixo_widget()

    st.markdown('<div class="titulo-secao">Efetivo padrão por banca</div>', unsafe_allow_html=True)
    capacidades = st.session_state["capacidade_bancas"]
    limites = st.session_state["limite_fora_sede_bancas"]
    df_bancas = pd.DataFrame(
        [
            {
                "Banca": banca,
                "Efetivo diário": int(capacidades.get(banca, 0) or 0),
                "Máximo fora da sede": int(limites.get(banca, 0) or 0),
            }
            for banca in bancas_config
        ]
    )
    editado = st.data_editor(
        df_bancas,
        hide_index=True,
        num_rows="fixed",
        **LARGURA_TOTAL,
        column_config={
            "Banca": st.column_config.TextColumn("Banca", disabled=True),
            "Efetivo diário": st.column_config.NumberColumn(
                "Efetivo diário", min_value=0, max_value=300, step=1, required=True,
                help="Examinadores da banca em um dia normal.",
            ),
            "Máximo fora da sede": st.column_config.NumberColumn(
                "Máximo fora da sede", min_value=0, max_value=300, step=1, required=True,
                help="Teto de examinadores em viagem ao mesmo tempo.",
            ),
        },
        key=f"efetivo_bancas_{assinatura_dados(df_bancas)[:8]}_{sufixo}",
    )
    mudou = False
    for registro in editado.to_dict(orient="records"):
        banca = registro["Banca"]
        efetivo = _inteiro(registro["Efetivo diário"])
        limite = _inteiro(registro["Máximo fora da sede"])
        if efetivo != int(capacidades.get(banca, 0) or 0):
            capacidades[banca] = efetivo
            mudou = True
        if limite != int(limites.get(banca, 0) or 0):
            limites[banca] = limite
            mudou = True
    if mudou:
        salvar_config("capacidade_bancas", capacidades)
        salvar_config("limite_fora_sede_bancas", limites)
        st.rerun()

    st.markdown('<div class="titulo-secao">Ajuste por dia</div>', unsafe_allow_html=True)
    st.caption(
        "Use para férias, licenças, cursos ou reforço em datas específicas. Dias sem"
        " ajuste seguem o efetivo padrão."
    )
    col_banca, col_mes, col_ano = st.columns([2, 1, 1])
    banca = col_banca.selectbox("Banca", list(bancas_config), key="ef_banca")
    mes_nome, ano, mes_numero = _seletor_mes_ano("ef", [col_mes, col_ano])

    dados = dados_da_banca_no_mes(banca, mes_nome, ano)
    totais = efetivo_diario_da_banca(banca, mes_nome, ano, dados)
    postos = postos_de_sobra(bancas_config, banca)
    sobra = sobra_do_mes(banca, mes_nome, ano, postos, dados) if postos else {}
    dias = [
        d
        for d in dias_do_intervalo(
            datetime.date(ano, mes_numero, 1),
            datetime.date(ano, mes_numero, calendar.monthrange(ano, mes_numero)[1]),
        )
        if d.weekday() < 6
    ]
    linhas = []
    for d in dias:
        em_campo = pico_diario(totais.get(d, {}).get("M", 0), totais.get(d, {}).get("T", 0))
        efetivo = efetivo_do_dia(banca, d)
        linha = {
            "Data": formatar_br(d),
            "Dia": nome_dia_semana(d),
            "Efetivo do dia": efetivo,
            "Em campo": em_campo,
            "Em viagem": totais.get(d, {}).get("viagem", 0),
            "Saldo": efetivo - em_campo,
        }
        if postos:
            linha["Livre p/ postos de sobra"] = min(sobra.get(d, (0, 0)))
        linhas.append(linha)
    df_dias = pd.DataFrame(linhas)

    def destacar(linha: pd.Series) -> list[str]:
        estilos = []
        for coluna, valor in linha.items():
            css = ""
            if coluna == "Saldo" and _inteiro(valor) < 0:
                css = f"background-color: {COR_ERRO_BG}; color: {COR_ERRO_TXT}; font-weight: 700;"
            if coluna == "Livre p/ postos de sobra" and _inteiro(valor) < MINIMO_SOBRA_POSTOS:
                css = f"background-color: {COR_ERRO_BG}; color: {COR_ERRO_TXT}; font-weight: 700;"
            if coluna == "Efetivo do dia" and _inteiro(valor) != efetivo_padrao_da_banca(banca):
                css = "background-color: #FEFCBF; font-weight: 700;"
            estilos.append(css)
        return estilos

    editado_dias = st.data_editor(
        df_dias.style.apply(destacar, axis=1),
        hide_index=True,
        num_rows="fixed",
        height=min(38 + 35 * len(df_dias), 760),
        **LARGURA_TOTAL,
        column_config={
            "Data": st.column_config.TextColumn("Data", disabled=True),
            "Dia": st.column_config.TextColumn("Dia", disabled=True),
            "Efetivo do dia": st.column_config.NumberColumn(
                "Efetivo do dia", min_value=0, max_value=300, step=1, required=True,
                help="Edite para registrar férias, licenças ou reforço nesta data.",
            ),
            "Em campo": st.column_config.NumberColumn("Em campo", disabled=True),
            "Em viagem": st.column_config.NumberColumn("Em viagem", disabled=True),
            "Saldo": st.column_config.NumberColumn("Saldo", disabled=True),
            "Livre p/ postos de sobra": st.column_config.NumberColumn(
                "Livre p/ postos de sobra", disabled=True,
                help=f"Mínimo de {MINIMO_SOBRA_POSTOS} nos dias de atendimento desses postos.",
            ),
        },
        key=f"efetivo_dias_{_assinatura(banca)[:8]}_{mes_numero}_{ano}_{assinatura_dados(df_dias)[:8]}_{sufixo}",
    )
    alterou = False
    for d, registro in zip(dias, editado_dias.to_dict(orient="records")):
        novo = _inteiro(registro["Efetivo do dia"])
        if novo != efetivo_do_dia(banca, d):
            definir_efetivo_do_dia(banca, d, novo)
            alterou = True
    if alterou:
        salvar_efetivo_dia(st.session_state["efetivo_dia_dict"])
        st.rerun()
    st.caption("Em amarelo: dias com ajuste. Em vermelho: saldo negativo ou sobra abaixo do mínimo.")


# --- Página: horários e regras -----------------------------------------------
PADRAO_HORARIO = re.compile(r"^([01]?\d|2[0-3]):([0-5]\d)$")


def _normalizar_horario(texto: str) -> str | None:
    bruto = (texto or "").strip().replace("h", ":").replace(".", ":")
    correspondencia = PADRAO_HORARIO.match(bruto)
    if not correspondencia:
        return None
    hora, minuto = correspondencia.groups()
    return f"{int(hora):02d}:{minuto}"


def _persistir_horarios(lista: list[str], inativos: list[str], fim_turno: list[str]) -> None:
    st.session_state["lista_horarios"] = lista
    st.session_state["horarios_inativos"] = inativos
    st.session_state["horarios_fim_turno"] = fim_turno
    salvar_config("lista_horarios", lista)
    salvar_config("horarios_inativos", inativos)
    salvar_config("horarios_fim_turno", fim_turno)


def _ordenar_horarios(horarios: Iterable[str]) -> list[str]:
    return sorted(set(horarios), key=lambda h: _hora_em_minutos(h) or 0)


def pagina_horarios() -> None:
    st.markdown("### ⏰ Horários, Turmas e Regras de Capacidade")
    st.info(
        "Os horários ativos geram as linhas de cada dia de atendimento. Horários"
        " desativados param de ser oferecidos; os lançamentos feitos neles ficam"
        " guardados, mas saem de todos os totais enquanto o horário estiver"
        " desativado."
    )

    lista: list[str] = list(st.session_state["lista_horarios"])
    inativos: list[str] = list(st.session_state["horarios_inativos"])
    fim_turno: list[str] = horarios_fim_turno()
    col_cadastro, col_visao = st.columns([1, 2])

    with col_cadastro:
        st.markdown("#### ➕ Adicionar Horário")
        novo = st.text_input("Horário (HH:MM):", placeholder="Ex: 17:30", key="hora_novo")
        if st.button("➕ Adicionar", key="hora_btn_add"):
            normalizado = _normalizar_horario(novo)
            if not normalizado:
                st.error("Formato inválido. Use HH:MM, por exemplo 14:30.")
            elif normalizado in lista:
                st.warning("Este horário já está cadastrado.")
            else:
                _persistir_horarios(_ordenar_horarios(lista + [normalizado]), inativos, fim_turno)
                _flash("success", f"Horário **{normalizado}** adicionado.")
                st.rerun()

        st.markdown("---")
        st.markdown("#### ♻️ Restaurar Grade Padrão")
        st.caption(
            "Volta aos oito horários originais, com 11:30 e 16:30 como fim de"
            " turno. Não altera calendários já lançados."
        )
        if st.button("♻️ Restaurar padrão", key="hora_btn_padrao"):
            _persistir_horarios(list(HORARIOS_PADRAO), [], list(HORARIOS_FIM_TURNO_PADRAO))
            _flash("success", "Grade padrão restaurada.")
            st.rerun()

    with col_visao:
        st.markdown("#### 🗂️ Horários Cadastrados")
        if not lista:
            st.warning("Nenhum horário cadastrado. Os calendários sairiam vazios.")
            return
        lista = _ordenar_horarios(lista)
        df_visao = pd.DataFrame(
            {
                "Horário": lista,
                "Período": [periodo_do_horario(h) for h in lista],
                "Ativo": [h not in inativos for h in lista],
                "Fim de turno (50%)": [eh_fim_de_turno(h, fim_turno) for h in lista],
            }
        )
        editado = st.data_editor(
            df_visao,
            hide_index=True,
            **LARGURA_TOTAL,
            num_rows="fixed",
            column_config={
                "Horário": st.column_config.TextColumn("Horário", disabled=True),
                "Período": st.column_config.TextColumn("Período", disabled=True),
                "Ativo": st.column_config.CheckboxColumn(
                    "Ativo", help="Desmarque para parar de oferecer este horário."
                ),
                "Fim de turno (50%)": st.column_config.CheckboxColumn(
                    "Fim de turno (50%)",
                    help="Nesses horários a capacidade de B, C, D e E cai pela metade.",
                ),
            },
            # A chave muda junto com a lista: as edições do widget são por
            # posição e não podem cair em outro horário depois de uma exclusão.
            key=f"hora_editor_{assinatura_dados(lista, inativos, fim_turno)[:10]}_{sufixo_widget()}",
        )
        novos_inativos = sorted(editado.loc[~editado["Ativo"].fillna(True).astype(bool), "Horário"].tolist())
        novos_fim = _ordenar_horarios(
            editado.loc[editado["Fim de turno (50%)"].fillna(False).astype(bool), "Horário"].tolist()
        )
        if novos_inativos != sorted(inativos) or novos_fim != _ordenar_horarios(fim_turno):
            _persistir_horarios(lista, novos_inativos, novos_fim)
            st.rerun()

        ativos = [h for h in lista if h not in inativos]
        st.markdown("---")
        col_m, col_t, col_total = st.columns(3)
        col_m.metric("Turmas de Manhã", sum(1 for h in ativos if periodo_do_horario(h) == TURNO_MANHA))
        col_t.metric("Turmas de Tarde", sum(1 for h in ativos if periodo_do_horario(h) == TURNO_TARDE))
        col_total.metric("Horários ativos", f"{len(ativos)} / {len(lista)}")

        st.markdown("#### 🗑️ Excluir Horário Definitivamente")
        alvo = st.selectbox("Horário:", lista, key="hora_excluir")
        st.caption("Para apenas suspender o horário, use a coluna *Ativo* acima.")
        if st.button("🗑️ Excluir", key="hora_btn_del"):
            _persistir_horarios(
                [h for h in lista if h != alvo],
                [h for h in inativos if h != alvo],
                [h for h in fim_turno if h != alvo],
            )
            _flash("warning", f"Horário **{alvo}** excluído da grade.")
            st.rerun()

    st.markdown("---")
    st.markdown("#### 📏 Regras de capacidade em vigor")
    st.dataframe(
        pd.DataFrame(
            [
                {"Categoria": "A (moto)", "Exames por examinador": 6, "Por dupla": 12,
                 "Fim de turno": "sem corte (metade no Pátio DETRAN)", "Turno": "só manhã",
                 "Horário quebrado": "-",
                 "Limite físico": "1 dupla por pista; blocos fechados de 12 por pista"},
                {"Categoria": "B", "Exames por examinador": 2, "Por dupla": 4,
                 "Fim de turno": "metade", "Turno": "manhã e tarde",
                 "Horário quebrado": "horário cheio", "Limite físico": "-"},
                {"Categoria": "C", "Exames por examinador": 2, "Por dupla": 4,
                 "Fim de turno": "metade", "Turno": "só manhã",
                 "Horário quebrado": "+1 min", "Limite físico": "-"},
                {"Categoria": "D", "Exames por examinador": 2, "Por dupla": 4,
                 "Fim de turno": "metade", "Turno": "só manhã",
                 "Horário quebrado": "+2 min", "Limite físico": "-"},
                {"Categoria": "E", "Exames por examinador": 2, "Por dupla": 4,
                 "Fim de turno": "metade", "Turno": "só manhã",
                 "Horário quebrado": "+10 min", "Limite físico": "-"},
            ]
        ),
        hide_index=True,
        **LARGURA_TOTAL,
    )
    st.caption(
        "Vagas PCD seguem a regra da categoria correspondente. A soma das cargas"
        " de todas as categorias no horário não pode passar dos examinadores"
        " escalados (Exam. M de manhã, Exam. T à tarde); a linha fica vermelha"
        " quando passa. Cat A: cada pista em uso oferta 12 vagas no horário (24 com"
        " 2 pistas e 4 examinadores); no Pátio do DETRAN, 11:30 oferta metade."
        f" Postos de sobra (Pátio DETRAN, Castelinho e Imperatriz): mínimo de"
        f" {MINIMO_SOBRA_POSTOS} examinadores livres por dia."
    )

    st.markdown("#### 🎭 Pontos facultativos que bloqueiam o calendário")
    ano = datetime.date.today().year
    nomes = sorted(
        {str(n) for a in (ano, ano + 1) for _, n in pontos_facultativos(a)}
        | set(st.session_state.get("facultativos_bloqueados", []))
    )
    if not nomes:
        st.caption("A versão instalada da biblioteca de feriados não informa pontos facultativos.")
        return
    escolhidos = st.multiselect(
        "Bloquear também nestes dias:",
        nomes,
        default=[n for n in st.session_state.get("facultativos_bloqueados", []) if n in nomes],
        key=f"facultativos_{sufixo_widget()}",
        help="Feriados nacionais e estaduais (MA) já bloqueiam sempre.",
    )
    if sorted(escolhidos) != sorted(st.session_state.get("facultativos_bloqueados", [])):
        st.session_state["facultativos_bloqueados"] = list(escolhidos)
        salvar_config("facultativos_bloqueados", list(escolhidos))
    proximos = [
        f"{formatar_br(d)} {n}"
        for a in (ano, ano + 1)
        for d, n in pontos_facultativos(a)
        if any(normalizar_texto(e) in normalizar_texto(n) for e in escolhidos)
        and d >= datetime.date.today()
    ]
    if proximos:
        st.caption("Próximos bloqueios: " + " · ".join(proximos[:8]))


# --- Página: dashboard consolidado -------------------------------------------
def _linha_resumo(
    banca: str, local: str, mes: str, ano: int, df_local: pd.DataFrame
) -> dict | None:
    df_vagas = df_local[(df_local["Status"] == STATUS_DISPONIVEL) & (df_local["Total"] > 0)]
    if df_vagas.empty:
        return None
    pico = max(
        pico_diario(*efetivo_do_grupo(g)) for g in registros_por_data(df_vagas).values()
    )
    dias = [d.day for d in (para_data(x) for x in df_vagas["Data"].unique()) if d]
    resumo = {
        "Banca": banca,
        "Localidade": local,
        "Mês/Ano": f"{mes}/{ano}",
        "Dias com Exame": int(df_vagas["Data"].nunique()),
        "Datas dos exames": formatar_datas_exames(dias),
        "Pico Exam./Dia": int(pico),
        "Horários com estouro": len(violacoes_do_local(banca, local, df_local)),
        "Total Vagas": int(df_vagas["Total"].fillna(0).sum()),
    }
    for coluna in COLS_VAGAS:
        resumo[coluna] = int(df_vagas[coluna].fillna(0).sum()) if coluna in df_vagas.columns else 0
    resumo[COL_DETALHE_PCD] = detalhe_pcd(df_vagas)
    return resumo


def ordenar_resumo_publicacao(df_resumo: pd.DataFrame) -> pd.DataFrame:
    """Linhas na ordem oficial da planilha de publicação."""
    if df_resumo.empty:
        return df_resumo
    bancas_config = st.session_state["bancas_config"]
    ordem = [
        chave_ordem_publicacao(b, l, bancas_config)
        for b, l in zip(df_resumo["Banca"], df_resumo["Localidade"])
    ]
    posicoes = sorted(range(len(ordem)), key=lambda i: ordem[i])
    return df_resumo.iloc[posicoes].reset_index(drop=True)


def pagina_relatorio() -> None:
    cabecalho_pagina(
        "Dashboard de Oferta de Vagas",
        "Consolidado do mês e planilha oficial de publicação.",
    )

    bancas_config = st.session_state["bancas_config"]
    col_mes, col_ano, col_banca, col_local = st.columns(4)
    mes_filtro, ano_filtro, _ = _seletor_mes_ano("rel", [col_mes, col_ano])
    banca_filtro = col_banca.selectbox(
        "Banca:", ["Todas"] + list(bancas_config.keys()), key="rel_banca"
    )
    opcoes_local = (
        bancas_config.get(banca_filtro, [])
        if banca_filtro != "Todas"
        else todas_localidades(bancas_config)
    )
    local_filtro = col_local.selectbox(
        "Localidade:", ["Todas as Localidades"] + opcoes_local, key="rel_local"
    )

    resumos = []
    bancas = list(bancas_config.keys()) if banca_filtro == "Todas" else [banca_filtro]
    for banca in bancas:
        for local, df_local in dados_da_banca_no_mes(banca, mes_filtro, ano_filtro).items():
            if local_filtro != "Todas as Localidades" and local != local_filtro:
                continue
            resumo = _linha_resumo(banca, local, mes_filtro, ano_filtro, df_local)
            if resumo:
                resumos.append(resumo)

    if not resumos:
        st.info("Nenhum dado cadastrado ou encontrado para os filtros selecionados.")
        return

    df_resumo = ordenar_resumo_publicacao(pd.DataFrame(resumos))

    total_vagas = int(df_resumo["Total Vagas"].sum())
    total_dias = int(df_resumo["Dias com Exame"].sum())
    media = round(total_vagas / total_dias, 1) if total_dias else 0
    pico_maximo = int(df_resumo["Pico Exam./Dia"].max())

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total de vagas ofertadas", f"{total_vagas:,}".replace(",", "."))
    kpi2.metric("Dias com exame (soma das localidades)", total_dias)
    kpi3.metric("Média de vagas por dia", f"{media}".replace(".", ","))
    kpi4.metric("Maior pico por localidade", f"{pico_maximo} exam.")

    estouros = int(df_resumo["Horários com estouro"].sum())
    if estouros:
        st.error(
            f"{estouros} horário(s) com capacidade estourada na seleção. Corrija"
            " na página de calendário antes de publicar."
        )

    total_categorias = int(df_resumo[COLS_CATEGORIA].sum().sum())
    total_pcd = int(df_resumo[COLS_PCD].sum().sum())
    if total_pcd:
        st.caption(
            f"Composição: {total_categorias} vagas de ampla concorrência e {total_pcd} vagas"
            " PCD. Na planilha oficial, as vagas PCD somam na categoria e ganham asterisco"
            " com a observação no fim."
        )

    por_categoria = pd.DataFrame(
        {
            "Categoria": [c.replace("Cat ", "") for c in COLS_CATEGORIA],
            "Vagas": [
                int(df_resumo[c].sum() + df_resumo[f"PCD {c.split()[1]}"].sum())
                for c in COLS_CATEGORIA
            ],
        }
    )
    col_grafico, col_tabela = st.columns([1, 2])
    with col_grafico:
        st.markdown('<div class="titulo-secao">Vagas por categoria</div>', unsafe_allow_html=True)
        st.bar_chart(por_categoria, x="Categoria", y="Vagas", color="#1D4E89", height=260)
    with col_tabela:
        st.markdown('<div class="titulo-secao">Vagas por banca</div>', unsafe_allow_html=True)
        por_banca = (
            df_resumo.groupby("Banca", sort=False)[["Total Vagas", "Dias com Exame"]]
            .sum()
            .reset_index()
        )
        st.dataframe(por_banca, hide_index=True, **LARGURA_TOTAL)

    st.markdown('<div class="titulo-secao">Panorama por localidade</div>', unsafe_allow_html=True)
    st.dataframe(
        df_resumo.drop(columns=[COL_DETALHE_PCD]), **LARGURA_TOTAL, hide_index=True
    )

    excel = gerar_excel_publicacao(df_resumo, mes_filtro, ano_filtro)
    st.download_button(
        label="Exportar relatório consolidado (Excel oficial)",
        data=excel,
        file_name=f"Calendario_Publicacao_{mes_filtro.upper()}_{ano_filtro}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="rel_download_excel",
        type="primary",
        icon=":material/download:",
    )
    st.caption(
        "Ordem da planilha: São Luís, Imperatriz, Timon, Bacabal, Caxias e Santa Inês."
        " Em São Luís, Pátio DETRAN, Área do Castelinho, Cohatrac, Cidade Operária, Paço"
        " do Lumiar, Raposa e São José de Ribamar, e depois os demais em ordem"
        " alfabética; nas outras bancas, o município da banca primeiro."
    )


# --- Página: gestão de bancas e localidades ----------------------------------
def _limpar_dados_da_localidade(banca: str, local: str) -> int:
    """Remove tudo que está pendurado em uma localidade. Devolve quantas
    viagens foram excluídas."""
    remover_historico_do_local(banca, local)
    chave = chave_local(banca, local)
    st.session_state["dias_permitidos_dict"].pop(chave, None)
    st.session_state["feriados_locais_dict"].pop(chave, None)
    st.session_state["parametros_local_dict"].pop(chave, None)

    restantes = [
        v
        for v in st.session_state["viagens_registradas"]
        if not (v.get("Banca") == banca and v.get("Destino") == local)
    ]
    removidas = len(st.session_state["viagens_registradas"]) - len(restantes)
    st.session_state["viagens_registradas"] = restantes

    salvar_dias_permitidos(st.session_state["dias_permitidos_dict"])
    salvar_feriados_locais(st.session_state["feriados_locais_dict"])
    salvar_parametros_local(st.session_state["parametros_local_dict"])
    if removidas:
        salvar_viagens(st.session_state["viagens_registradas"])
    LOG.info("Localidade %s/%s removida (%d viagens excluídas)", banca, local, removidas)
    return removidas


def _contar_dados_da_banca(banca: str) -> tuple[int, int]:
    calendarios = sum(
        1
        for chave in st.session_state["historico_localidades"]
        if (partes := desmontar_chave_historico(chave)) and partes[0] == banca
    )
    viagens = sum(1 for v in st.session_state["viagens_registradas"] if v.get("Banca") == banca)
    return calendarios, viagens


def _remover_banca(banca: str) -> None:
    bancas_config = st.session_state["bancas_config"]
    for local in list(bancas_config.get(banca, [])):
        _limpar_dados_da_localidade(banca, local)
    # Viagens da banca que sobraram e participação dela como apoio em outras.
    viagens = [v for v in st.session_state["viagens_registradas"] if v.get("Banca") != banca]
    for viagem in viagens:
        if banca in (viagem.get("Bancas Apoio") or []):
            viagem["Bancas Apoio"] = [b for b in viagem["Bancas Apoio"] if b != banca]
            por_banca = dict(viagem.get("Examinadores Por Banca") or {})
            por_banca.pop(banca, None)
            viagem["Examinadores Por Banca"] = por_banca
            viagem["Examinadores"] = total_examinadores(viagem)
    st.session_state["viagens_registradas"] = viagens
    salvar_viagens(viagens)

    disponibilidade = st.session_state["disponibilidade_semanal"]
    for chave in [c for c in disponibilidade if c.startswith(_sanitizar(banca) + SEPARADOR_CHAVE)]:
        disponibilidade.pop(chave, None)
    salvar_disponibilidade(disponibilidade)
    ajustes = st.session_state["efetivo_dia_dict"]
    for chave in [c for c in ajustes if c.startswith(_sanitizar(banca) + SEPARADOR_CHAVE)]:
        ajustes.pop(chave, None)
    salvar_efetivo_dia(ajustes)

    bancas_config.pop(banca, None)
    st.session_state["capacidade_bancas"].pop(banca, None)
    st.session_state["limite_fora_sede_bancas"].pop(banca, None)
    salvar_bancas(bancas_config)
    salvar_config("capacidade_bancas", st.session_state["capacidade_bancas"])
    salvar_config("limite_fora_sede_bancas", st.session_state["limite_fora_sede_bancas"])


def pagina_gestao() -> None:
    cabecalho_pagina(
        "Bancas e Localidades",
        "Cadastro de bancas e municípios. A remoção apaga também os calendários,"
        " feriados, parâmetros e viagens vinculados (um snapshot de segurança é"
        " gravado antes).",
    )

    bancas_config: dict[str, list[str]] = st.session_state["bancas_config"]
    col_banca, col_local = st.columns(2)

    with col_banca:
        st.markdown("#### 🏛️ Cadastrar Nova Banca")
        nova_banca = st.text_input(
            "Nome da Nova Banca:", placeholder="Ex: Banca Balsas", key="gestao_nova_banca"
        )
        if st.button("➕ Adicionar Banca", key="gestao_btn_add_banca"):
            nome = (nova_banca or "").strip()
            if not nome:
                st.warning("Informe um nome para a banca.")
            elif SEPARADOR_CHAVE in nome:
                st.warning(f"O nome não pode conter '{SEPARADOR_CHAVE}'.")
            elif nome in bancas_config:
                st.warning("Esta banca já existe no sistema.")
            else:
                bancas_config[nome] = []
                salvar_bancas(bancas_config)
                _flash("success", f"Banca **{nome}** criada com sucesso!")
                st.rerun()

    with col_local:
        st.markdown("#### 📍 Cadastrar Novo Município / Localidade")
        if not bancas_config:
            st.warning("Cadastre uma banca antes de adicionar localidades.")
        else:
            banca_destino = st.selectbox(
                "Selecione a Banca que receberá o local:",
                list(bancas_config.keys()),
                key="gestao_banca_destino",
            )
            novo_local = st.text_input(
                "Nome do Município/Localidade:", placeholder="Ex: Balsas - Pátio", key="gestao_novo_local"
            )
            st.caption(
                "A localidade entra na ordem oficial (sede ou região metropolitana"
                " primeiro, o resto em ordem alfabética). Fora da região"
                " metropolitana de São Luís e de Imperatriz, e fora da sede nas"
                " demais bancas, o município é atendido por viagem itinerante."
            )
            if st.button("➕ Adicionar Localidade", key="gestao_btn_add_local"):
                nome = (novo_local or "").strip()
                if not nome:
                    st.warning("Informe um nome para a localidade.")
                elif SEPARADOR_CHAVE in nome:
                    st.warning(f"O nome não pode conter '{SEPARADOR_CHAVE}'.")
                elif nome in bancas_config[banca_destino]:
                    st.warning("Esta localidade já está cadastrada nesta banca.")
                else:
                    bancas_config[banca_destino] = ordenar_localidades_da_banca(
                        banca_destino, bancas_config[banca_destino] + [nome]
                    )
                    salvar_bancas(bancas_config)
                    _flash("success", f"Localidade **{nome}** vinculada à **{banca_destino}**!")
                    st.rerun()

    st.markdown("---")
    st.markdown("#### 🗑️ Remover Localidade Existente")
    if not bancas_config:
        st.info("Nenhuma banca cadastrada.")
        return
    col_banca_rem, col_local_rem = st.columns(2)
    with col_banca_rem:
        banca_remocao = st.selectbox(
            "Selecione a Banca:", list(bancas_config.keys()), key="gestao_banca_remocao"
        )
    with col_local_rem:
        locais = bancas_config.get(banca_remocao, [])
        if not locais:
            st.info(f"A banca **{banca_remocao}** não possui localidades cadastradas.")
        else:
            local_remocao = st.selectbox(
                "Selecione a Localidade para Remover:", locais, key="gestao_local_remocao"
            )
            confirmar = st.checkbox(
                f"Confirmo a exclusão de **{local_remocao}** e de todos os dados vinculados a ela.",
                key="gestao_confirma_local",
            )
            if st.button("🗑️ Remover Localidade", key="gestao_btn_rem_local"):
                if not confirmar:
                    st.warning("Marque a confirmação antes de remover.")
                else:
                    salvar_snapshot(motivo="antes_remover_localidade")
                    _limpar_dados_da_localidade(banca_remocao, local_remocao)
                    bancas_config[banca_remocao].remove(local_remocao)
                    salvar_bancas(bancas_config)
                    _flash("warning", f"Localidade **{local_remocao}** e seus dados foram removidos.")
                    st.rerun()

    st.markdown("---")
    with st.expander("⚠️ Remover uma Banca inteira", expanded=False):
        banca_excluir = st.selectbox(
            "Banca a excluir:", list(bancas_config.keys()), key="gestao_banca_excluir"
        )
        calendarios, viagens = _contar_dados_da_banca(banca_excluir)
        st.write(
            f"Serão apagados: **{len(bancas_config.get(banca_excluir, []))}**"
            f" localidades, **{calendarios}** calendários e **{viagens}** viagens."
        )
        confirmar_banca = st.checkbox(
            "Confirmo a exclusão definitiva desta banca.", key="gestao_confirma_banca"
        )
        if st.button("🗑️ Excluir Banca", key="gestao_btn_rem_banca"):
            if not confirmar_banca:
                st.warning("Marque a confirmação antes de excluir.")
            else:
                salvar_snapshot(motivo="antes_remover_banca")
                _remover_banca(banca_excluir)
                _flash("warning", f"Banca **{banca_excluir}** excluída.")
                st.rerun()


# ===========================================================================
# 7. APLICAÇÃO
# ===========================================================================
def configurar_logging() -> None:
    nivel = os.environ.get("DETRAN_LOG_LEVEL", "INFO").upper()
    logging.basicConfig(
        level=getattr(logging, nivel, logging.INFO),
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    )


def _senha_configurada() -> str | None:
    senha = os.environ.get("DETRAN_SENHA")
    if senha:
        return senha
    try:
        return st.secrets.get("senha")  # type: ignore[no-any-return]
    except Exception:
        return None


def exigir_senha() -> bool:
    """Bloqueio simples por senha, ligado só se DETRAN_SENHA (ou
    st.secrets["senha"]) estiver definida. Sem isso, qualquer pessoa com o
    link podia apagar bancas."""
    senha = _senha_configurada()
    if not senha or st.session_state.get("_autenticado"):
        return True
    st.markdown("### 🔒 Acesso restrito")
    digitada = st.text_input("Senha:", type="password", key="campo_senha")
    if st.button("Entrar", key="btn_entrar"):
        if hmac.compare_digest(str(digitada), str(senha)):
            st.session_state["_autenticado"] = True
            st.rerun()
        st.error("Senha incorreta.")
    return False


def painel_backup() -> None:
    st.sidebar.markdown("---")
    with st.sidebar.expander("Backup e sincronização", expanded=False, icon=":material/backup:"):
        st.caption(
            "Os dados ficam em banco SQLite e são gravados assim que você altera"
            " algo. Use os botões abaixo para guardar uma cópia externa ou"
            " restaurar um backup."
        )
        nome_arquivo = f"backup_detran_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.json"
        if st.button("📦 Preparar backup para download", key="btn_preparar_backup"):
            st.session_state["_backup_pronto"] = backup_em_json()
        if st.session_state.get("_backup_pronto"):
            st.download_button(
                "⬇️ Baixar backup",
                data=st.session_state["_backup_pronto"],
                file_name=nome_arquivo,
                mime="application/json",
                key="btn_baixar_backup",
            )

        if st.button("🗂️ Gravar snapshot no servidor", key="btn_snapshot"):
            destino = salvar_snapshot(motivo="manual")
            if destino:
                st.success(f"Snapshot gravado: `{destino.name}`")

        st.markdown("---")
        arquivo = st.file_uploader("Restaurar um backup (.json):", type=["json"], key="upload_backup")
        if arquivo is not None:
            confirmar = st.checkbox(
                "Confirmo que este backup deve substituir TODOS os dados atuais.",
                key="confirma_restauracao",
            )
            if st.button("♻️ Restaurar este backup", key="btn_restaurar"):
                if not confirmar:
                    st.warning("Marque a confirmação antes de restaurar.")
                else:
                    try:
                        aplicar_backup(json.load(arquivo))
                    except BackupInvalido as erro:
                        st.error(f"Backup recusado: {erro}")
                    except json.JSONDecodeError as erro:
                        st.error(f"O arquivo não é um JSON válido: {erro}")
                    except Exception as erro:
                        LOG.exception("Falha ao restaurar backup")
                        st.error(f"Falha ao restaurar (nada foi alterado): {erro}")
                    else:
                        _flash("success", "Backup restaurado com sucesso!")
                        st.rerun()

        st.markdown("---")
        st.caption(
            "Se outra pessoa estiver usando o sistema ao mesmo tempo, recarregue"
            " para ver as alterações dela. As suas gravações não apagam as dela."
        )
        if st.button("🔄 Recarregar dados do banco", key="btn_recarregar"):
            recarregar_do_banco()
            st.rerun()


def barra_de_status() -> None:
    erro = st.session_state.get("erro_persistencia")
    if erro:
        st.error(
            f"O salvamento apresentou erro e os dados podem não ter sido gravados: {erro}",
            icon=":material/error:",
        )
    elif st.session_state.get("ultimo_salvamento"):
        st.sidebar.caption(f"Última gravação: {st.session_state['ultimo_salvamento']}")


PAGINA_QUADRO = "Quadro Geral"
PAGINA_CALENDARIO = "Montar Calendário"
PAGINA_VIAGENS = "Viagens"
PAGINA_EFETIVO = "Efetivo"
PAGINA_HORARIOS = "Horários & Regras"
PAGINA_DASHBOARD = "Dashboard"
PAGINA_LOCALIDADES = "Localidades"

PAGINAS = {
    PAGINA_QUADRO: pagina_quadro,
    PAGINA_CALENDARIO: pagina_calendario,
    PAGINA_VIAGENS: pagina_viagens,
    PAGINA_EFETIVO: pagina_efetivo,
    PAGINA_HORARIOS: pagina_horarios,
    PAGINA_DASHBOARD: pagina_relatorio,
    PAGINA_LOCALIDADES: pagina_gestao,
}
ICONES_PAGINA = {
    PAGINA_QUADRO: ":material/grid_view:",
    PAGINA_CALENDARIO: ":material/edit_calendar:",
    PAGINA_VIAGENS: ":material/directions_bus:",
    PAGINA_EFETIVO: ":material/groups:",
    PAGINA_HORARIOS: ":material/schedule:",
    PAGINA_DASHBOARD: ":material/monitoring:",
    PAGINA_LOCALIDADES: ":material/location_city:",
}


def seletor_de_pagina() -> str:
    """Navegação que desenha só a página escolhida.

    Com ``st.tabs`` todas as abas rodavam a cada clique (lento, e os
    parâmetros do calendário apareciam na barra lateral de todas as abas).
    """
    opcoes = list(PAGINAS)
    if st.session_state.get("pagina") not in (None, *opcoes):
        st.session_state["pagina"] = opcoes[0]
    if hasattr(st, "segmented_control"):
        escolha = st.segmented_control(
            "Página",
            opcoes,
            default=opcoes[0],
            key="pagina",
            label_visibility="collapsed",
            format_func=lambda p: f"{ICONES_PAGINA.get(p, '')} {p}",
        )
    else:
        escolha = st.radio(
            "Página", opcoes, horizontal=True, key="pagina", label_visibility="collapsed"
        )
    return escolha or opcoes[0]


def main() -> None:
    configurar_logging()
    st.set_page_config(
        page_title="DETRAN/MA - Gestão de Exames Práticos",
        page_icon=":material/directions_car:",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)
    st.markdown(CSS_CUSTOMIZADO, unsafe_allow_html=True)

    if not exigir_senha():
        return

    try:
        carregar_estado()
    except Exception as erro:
        LOG.exception("Falha ao carregar o estado inicial do banco de dados")
        st.error(
            "🚨 Não foi possível abrir o banco de dados do sistema. Nenhum dado"
            f" foi lido ou alterado. Detalhe técnico: {erro}"
        )
        st.info(
            "Verifique se o arquivo do banco (variável `DETRAN_DB`) existe e"
            " tem permissão de leitura/escrita, e tente recarregar a página."
        )
        st.stop()

    st.markdown(cabecalho_html(), unsafe_allow_html=True)
    barra_de_status()

    if st.session_state.get("modo_foco_calendario"):
        mostrar_flash()
        _renderizar_modo_foco_calendario()
    else:
        pagina = seletor_de_pagina()
        mostrar_flash()
        PAGINAS[pagina]()
    # Backup no fim da barra lateral, abaixo dos parâmetros da página.
    painel_backup()


if __name__ == "__main__":
    main()


