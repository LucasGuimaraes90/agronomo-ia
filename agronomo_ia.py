import streamlit as st
import anthropic
import base64
import io
import re
from datetime import datetime
import json

# ── PDF ──────────────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, HRFlowable, KeepTogether
)

# ── Excel ─────────────────────────────────────────────────────────────────────
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Agrônomo IA — Guara Agro",
    page_icon="🌱",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── CSS tema verde claro / white ───────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* fundo geral — branco */
.stApp { background: #ffffff; color: #1a2e1a; }

/* sidebar — verde muito suave */
section[data-testid="stSidebar"] { background: #f1f8f1; border-right: 1px solid #c8e6c9; }

/* chat bubbles — usuário */
[data-testid="stChatMessage"][data-testid*="user"] {
    background: #e8f5e9;
    border-radius: 12px;
    margin: 6px 0;
    border: 1px solid #c8e6c9;
}

/* chat bubbles — assistente */
[data-testid="stChatMessage"] {
    background: #f9fdf9;
    border-radius: 12px;
    margin: 6px 0;
    border: 1px solid #e0f0e0;
}

/* texto dentro dos bubbles */
[data-testid="stChatMessage"] p,
[data-testid="stChatMessage"] li,
[data-testid="stChatMessage"] td,
[data-testid="stChatMessage"] th { color: #1a2e1a !important; }

/* tabelas no chat */
[data-testid="stChatMessage"] table {
    border-collapse: collapse;
    width: 100%;
    background: #ffffff;
}
[data-testid="stChatMessage"] th {
    background: #2e7d32 !important;
    color: white !important;
    padding: 6px 10px;
}
[data-testid="stChatMessage"] td { padding: 5px 10px; border-bottom: 1px solid #e0f0e0; }
[data-testid="stChatMessage"] tr:nth-child(even) td { background: #f1f8f1; }

/* input de chat */
[data-testid="stChatInput"] textarea {
    background: #f9fdf9 !important;
    color: #1a2e1a !important;
    border: 1.5px solid #2e7d32 !important;
    border-radius: 10px !important;
}

/* botões primários */
.stButton > button {
    background: linear-gradient(135deg, #2e7d32, #1b5e20);
    color: white;
    border: none;
    border-radius: 8px;
    font-weight: 600;
    padding: 0.5rem 1.2rem;
    transition: all .2s;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #388e3c, #2e7d32);
    transform: translateY(-1px);
}

/* download buttons */
a.download-btn {
    display: inline-block;
    background: linear-gradient(135deg, #1565c0, #0d47a1);
    color: white !important;
    border-radius: 8px;
    padding: 0.45rem 1.1rem;
    text-decoration: none;
    font-size: 0.85rem;
    font-weight: 600;
    margin-right: 8px;
    margin-top: 6px;
}
a.download-btn:hover { background: linear-gradient(135deg, #1976d2, #1565c0); }
a.download-btn.green { background: linear-gradient(135deg, #2e7d32, #1b5e20); }

/* título principal */
h1 { color: #2e7d32 !important; font-weight: 900; }
h2, h3 { color: #388e3c !important; }

/* file uploader */
[data-testid="stFileUploader"] {
    background: #f9fdf9;
    border: 1.5px dashed #2e7d32;
    border-radius: 10px;
}

/* divider */
hr { border-color: #c8e6c9 !important; }

/* expander */
details { background: #f9fdf9; border-radius: 8px; border: 1px solid #c8e6c9; }

/* disclaimer box */
.disclaimer-box {
    background: #fff8e1;
    border: 1.5px solid #f9a825;
    border-radius: 8px;
    padding: 10px 14px;
    color: #5d4037 !important;
    font-size: 0.88rem;
    margin-top: 12px;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SYSTEM PROMPT
# ══════════════════════════════════════════════════════════════════════════════
SYSTEM_PROMPT = """Você é o **Agrônomo IA**, assistente técnico agrícola especializado da Guara Agro.
Foi treinado com as principais referências brasileiras de fertilidade do solo: 5ª Aproximação CFSEMG (1999), EMBRAPA, Apostila de Nutrição de Plantas (UFV/UNIPAM).

**Propósito:** ajudar produtores rurais, agrônomos e técnicos a interpretar análises de solo, calcular calagem e adubação, diagnosticar deficiências nutricionais e recomendar manejos culturais.

## REGRAS DE COMPORTAMENTO
- Sempre em português do Brasil, linguagem clara para o produtor rural
- Responda de forma direta e objetiva
- Use emojis com moderação (🌱 ✅ ⚠️)
- Nunca invente dados — se não sabe, diz claramente
- Este é um agente WEB — pode usar markdown completo (tabelas, negrito, listas)

## QUANDO FALTAM DADOS
Antes de calagem ou adubação, pergunte:
1. Textura do solo (% argila ou tipo: arenoso/médio/argiloso)
2. Cultura e produtividade esperada
3. Resultados da análise (pH, MO, P, K, Ca, Mg, Al, H+Al, CTC, V%)
4. PRNT do calcário
5. Sistema de plantio

Se o usuário enviar os dados de uma vez, calcule direto.

## SEQUÊNCIA AO RECEBER LAUDO (foto ou texto)
1. Interprete cada parâmetro (adequado / baixo / alto / excessivo)
2. Calcule NC pelos métodos da 5ª Aprox e apresente o mais conservador
3. Verifique necessidade de gessagem (Al³⁺ > 0,5 ou V% < 50 em subsolo)
4. Recomende NPK para a cultura/produtividade informada
5. Alerte sobre micronutrientes críticos (B, Zn, Cu, Mn)
6. Resumo final com tabela de ações práticas

## FORMULADOS — REGRA MÚLTIPLAS OPÇÕES
Ao recomendar adubo, SEMPRE dê 2 a 4 alternativas:

| Nutriente | Opção 1 | Opção 2 | Opção 3 |
|-----------|---------|---------|---------|
| P₂O₅ | MAP 09-46-00 | 10-40-10 | TSP 00-46-00 |
| N | Ureia 45-00-00 | Sulfato amônio 21-00-00 | Nitrato amônio 33-00-00 |
| K | KCl 00-00-60 | Sulfato K 00-00-50 | KCl granulado 00-00-58 |

Formulados completos cerrado mineiro: 08-28-16, 10-20-20, 04-30-10, 05-25-15

## TABELAS
Sempre que tiver dados comparativos ou recomendações, apresente em tabela markdown.
Exemplo de tabela de resultado:

| Parâmetro | Resultado | Referência | Status |
|-----------|-----------|------------|--------|
| pH H₂O | 5,1 | 5,5 – 6,5 | ⚠️ Baixo |

## GERAÇÃO DE ARQUIVOS
Quando o usuário pedir PDF, relatório, planilha ou Excel, avise que pode usar os botões de download abaixo do chat para baixar a última resposta como PDF ou exportar os dados como Excel.

## AVISO FINAL — SEMPRE INCLUIR
⚠️ *Esta recomendação não substitui o laudo de Engenheiro Agrônomo habilitado (CREA).*
"""

# ══════════════════════════════════════════════════════════════════════════════
# AUTENTICAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
VALID_CODES = st.secrets.get("ACCESS_CODES", "DEMO2024,GUARA2024,AGRO2024").split(",")

def check_auth():
    if st.session_state.get("authenticated"):
        return True
    st.markdown("## 🌱 Agrônomo IA — Guara Agro")
    st.markdown("### Acesso ao Agente")
    code = st.text_input("Código de acesso:", type="password", placeholder="Digite seu código...")
    if st.button("Entrar", use_container_width=True):
        if code.strip().upper() in [c.strip().upper() for c in VALID_CODES]:
            st.session_state.authenticated = True
            st.session_state.user_code = code.strip().upper()
            st.rerun()
        else:
            st.error("❌ Código inválido. Verifique com a Guara Agro.")
    st.markdown("---")
    st.caption("Ainda não tem acesso? Entre em contato via WhatsApp.")
    return False

# ══════════════════════════════════════════════════════════════════════════════
# GERAÇÃO DE PDF
# ══════════════════════════════════════════════════════════════════════════════
def markdown_to_pdf(markdown_text: str, title: str = "Relatório — Agrônomo IA") -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2.5*cm, bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    verde_escuro = colors.HexColor("#1b5e20")
    verde_medio  = colors.HexColor("#2e7d32")
    verde_claro  = colors.HexColor("#e8f5e9")
    cinza_bg     = colors.HexColor("#f5f5f5")

    estilo_titulo = ParagraphStyle(
        "Titulo", parent=styles["Title"],
        textColor=verde_escuro, fontSize=20, spaceAfter=6,
        fontName="Helvetica-Bold"
    )
    estilo_h2 = ParagraphStyle(
        "H2", parent=styles["Heading2"],
        textColor=verde_medio, fontSize=14, spaceBefore=14, spaceAfter=4,
        fontName="Helvetica-Bold"
    )
    estilo_h3 = ParagraphStyle(
        "H3", parent=styles["Heading3"],
        textColor=verde_escuro, fontSize=12, spaceBefore=10, spaceAfter=3,
        fontName="Helvetica-Bold"
    )
    estilo_body = ParagraphStyle(
        "Body", parent=styles["Normal"],
        fontSize=10, leading=15, spaceAfter=4,
        fontName="Helvetica"
    )
    estilo_aviso = ParagraphStyle(
        "Aviso", parent=styles["Normal"],
        fontSize=9, textColor=colors.HexColor("#b71c1c"),
        borderPad=4, backColor=colors.HexColor("#fff3e0"),
        borderColor=colors.HexColor("#e65100"), borderWidth=0.5,
        fontName="Helvetica-Oblique", leading=13
    )

    story = []

    # Cabeçalho
    story.append(Paragraph("🌱 GUARA AGRO", ParagraphStyle(
        "Logo", parent=styles["Normal"],
        textColor=verde_medio, fontSize=11, fontName="Helvetica-Bold"
    )))
    story.append(Spacer(1, 4))
    story.append(Paragraph(title, estilo_titulo))
    story.append(Paragraph(
        f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')} · Agrônomo IA v1.0",
        ParagraphStyle("sub", parent=styles["Normal"], fontSize=9,
                       textColor=colors.grey, fontName="Helvetica")
    ))
    story.append(HRFlowable(width="100%", thickness=2, color=verde_medio, spaceAfter=12))

    # Parser linha a linha
    lines = markdown_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]

        # Tabelas markdown
        if "|" in line and i + 1 < len(lines) and "|" in lines[i+1] and "---" in lines[i+1]:
            table_lines = []
            while i < len(lines) and "|" in lines[i]:
                cols = [c.strip() for c in lines[i].split("|") if c.strip()]
                if "---" not in lines[i]:
                    table_lines.append(cols)
                i += 1
            if table_lines:
                max_cols = max(len(r) for r in table_lines)
                data = [r + [""] * (max_cols - len(r)) for r in table_lines]
                col_width = (A4[0] - 4*cm) / max_cols
                t = Table(data, colWidths=[col_width] * max_cols)
                t.setStyle(TableStyle([
                    ("BACKGROUND",  (0, 0), (-1, 0), verde_medio),
                    ("TEXTCOLOR",   (0, 0), (-1, 0), colors.white),
                    ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE",    (0, 0), (-1, -1), 9),
                    ("BACKGROUND",  (0, 1), (-1, -1), cinza_bg),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, cinza_bg]),
                    ("GRID",        (0, 0), (-1, -1), 0.5, colors.HexColor("#c8e6c9")),
                    ("ALIGN",       (0, 0), (-1, -1), "LEFT"),
                    ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
                    ("PADDING",     (0, 0), (-1, -1), 5),
                    ("TOPPADDING",  (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]))
                story.append(t)
                story.append(Spacer(1, 8))
            continue

        # Títulos
        if line.startswith("### "):
            story.append(Paragraph(line[4:].strip(), estilo_h3))
        elif line.startswith("## "):
            story.append(Paragraph(line[3:].strip(), estilo_h2))
        elif line.startswith("# "):
            story.append(Paragraph(line[2:].strip(), estilo_titulo))

        # Listas
        elif line.strip().startswith(("- ", "* ", "• ")):
            txt = line.strip()[2:].strip()
            txt = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', txt)
            story.append(Paragraph(f"• {txt}", ParagraphStyle(
                "li", parent=estilo_body, leftIndent=14, spaceAfter=2
            )))
        elif re.match(r'^\d+\.', line.strip()):
            txt = line.strip()
            txt = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', txt)
            story.append(Paragraph(txt, ParagraphStyle(
                "ol", parent=estilo_body, leftIndent=14, spaceAfter=2
            )))

        # Aviso ⚠️
        elif "⚠️" in line or "não substitui" in line.lower():
            clean = re.sub(r'\*(.+?)\*', r'<i>\1</i>', line.strip().lstrip("⚠️ *"))
            story.append(Spacer(1, 6))
            story.append(Paragraph(f"⚠️ {clean}", estilo_aviso))

        # HR
        elif line.strip() in ("---", "___", "***"):
            story.append(HRFlowable(width="100%", thickness=0.5,
                                     color=colors.HexColor("#c8e6c9"), spaceAfter=6))

        # Linha vazia
        elif line.strip() == "":
            story.append(Spacer(1, 4))

        # Parágrafo normal
        else:
            txt = line.strip()
            txt = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', txt)
            txt = re.sub(r'\*(.+?)\*',   r'<i>\1</i>', txt)
            txt = re.sub(r'`(.+?)`',     r'<font name="Courier">\1</font>', txt)
            if txt:
                story.append(Paragraph(txt, estilo_body))

        i += 1

    # Rodapé
    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=1, color=verde_medio))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        "Guara Agro · Agrônomo IA v1.0 · ⚠️ Esta recomendação não substitui o laudo de Engenheiro Agrônomo habilitado (CREA).",
        ParagraphStyle("footer", parent=styles["Normal"], fontSize=8,
                       textColor=colors.grey, alignment=1, fontName="Helvetica-Oblique")
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer.read()


# ══════════════════════════════════════════════════════════════════════════════
# GERAÇÃO DE EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def extract_tables_from_markdown(text: str):
    """Extrai tabelas markdown e retorna lista de (header, rows)."""
    tables = []
    lines = text.split("\n")
    i = 0
    while i < len(lines):
        if "|" in lines[i]:
            raw = []
            while i < len(lines) and "|" in lines[i]:
                cols = [c.strip() for c in lines[i].split("|") if c.strip()]
                if "---" not in lines[i]:
                    raw.append(cols)
                i += 1
            if len(raw) >= 2:
                tables.append((raw[0], raw[1:]))
        else:
            i += 1
    return tables


def response_to_excel(markdown_text: str, title: str = "Agrônomo IA") -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Recomendações"

    verde = "1B5E20"
    verde_claro = "E8F5E9"
    verde_medio = "2E7D32"
    amarelo = "FFF9C4"

    # Cabeçalho
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = f"🌱 GUARA AGRO — {title}"
    c.font = Font(bold=True, size=16, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=verde_medio)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:G2")
    c2 = ws["A2"]
    c2.value = f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} · Agrônomo IA v1.0"
    c2.font = Font(italic=True, size=9, color="666666")
    c2.alignment = Alignment(horizontal="center")

    row = 4

    tables = extract_tables_from_markdown(markdown_text)

    if tables:
        for header, rows in tables:
            # Header da tabela
            for col_idx, h in enumerate(header, 1):
                cell = ws.cell(row=row, column=col_idx, value=h)
                cell.font = Font(bold=True, color="FFFFFF", size=11)
                cell.fill = PatternFill("solid", fgColor=verde)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(
                    bottom=Side(style="medium", color=verde_medio),
                    right=Side(style="thin", color="CCCCCC")
                )
                ws.column_dimensions[cell.column_letter].width = max(20, len(str(h)) + 4)
            ws.row_dimensions[row].height = 22
            row += 1

            # Linhas
            for r_idx, data_row in enumerate(rows):
                bg = verde_claro if r_idx % 2 == 0 else "FFFFFF"
                for col_idx, val in enumerate(data_row, 1):
                    cell = ws.cell(row=row, column=col_idx, value=val)
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.border = Border(
                        bottom=Side(style="thin", color="E0E0E0"),
                        right=Side(style="thin", color="E0E0E0")
                    )
                    cell.font = Font(size=10)
                ws.row_dimensions[row].height = 18
                row += 1
            row += 1

    else:
        # Sem tabela — dump do texto
        linhas = [l for l in markdown_text.split("\n") if l.strip()]
        ws.cell(row=row, column=1, value="Análise / Recomendação").font = Font(bold=True)
        row += 1
        for linha in linhas:
            ws.cell(row=row, column=1, value=linha)
            row += 1

    # Rodapé
    row += 1
    ws.merge_cells(f"A{row}:G{row}")
    c_footer = ws[f"A{row}"]
    c_footer.value = "⚠️ Esta recomendação não substitui o laudo de Engenheiro Agrônomo habilitado (CREA). · Guara Agro"
    c_footer.font = Font(italic=True, size=8, color="999999")
    c_footer.alignment = Alignment(horizontal="center")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
# HELPER: download link HTML
# ══════════════════════════════════════════════════════════════════════════════
def make_download_link(data: bytes, filename: str, label: str, css_class: str = "") -> str:
    b64 = base64.b64encode(data).decode()
    mime = "application/pdf" if filename.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return (
        f'<a href="data:{mime};base64,{b64}" download="{filename}" '
        f'class="download-btn {css_class}">{label}</a>'
    )


# ══════════════════════════════════════════════════════════════════════════════
# CLAUDE API
# ══════════════════════════════════════════════════════════════════════════════
def get_client():
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("ANTHROPIC_API_KEY não configurada em .streamlit/secrets.toml")
        st.stop()
    return anthropic.Anthropic(api_key=api_key)


def call_claude(messages: list, image_b64: str = None, image_mime: str = None) -> str:
    client = get_client()

    api_messages = []
    for m in messages[:-1]:  # histórico
        api_messages.append({"role": m["role"], "content": m["content"]})

    # Última mensagem (pode ter imagem)
    last = messages[-1]
    if image_b64 and last["role"] == "user":
        content = [
            {"type": "image", "source": {"type": "base64", "media_type": image_mime, "data": image_b64}},
            {"type": "text", "text": last["content"]},
        ]
    else:
        content = last["content"]

    api_messages.append({"role": "user", "content": content})

    with st.spinner("🌱 Analisando..."):
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=2048,
            system=SYSTEM_PROMPT,
            messages=api_messages,
        )
    return response.content[0].text


# ══════════════════════════════════════════════════════════════════════════════
# INTERFACE PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def main():
    # Header
    col1, col2 = st.columns([5, 2])
    with col1:
        st.markdown("# 🌱 Agrônomo IA")
        st.caption("Guara Agro · Assistente técnico especializado em solo e nutrição")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Sair", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    st.divider()

    # Inicializa histórico
    if "messages" not in st.session_state:
        st.session_state.messages = []
        st.session_state.messages.append({
            "role": "assistant",
            "content": (
                "👋 Olá! Sou o **Agrônomo IA** da Guara Agro.\n\n"
                "Posso te ajudar com:\n"
                "- 📊 **Interpretação de laudos de solo**\n"
                "- 🧮 **Cálculo de calagem e adubação NPK**\n"
                "- 🔍 **Diagnóstico de deficiências nutricionais**\n"
                "- 🌿 **Recomendações para o Cerrado Mineiro**\n\n"
                "Envie o laudo (foto ou texto), faça uma pergunta ou me peça para gerar um relatório. "
                "Depois de cada análise, você poderá **baixar o relatório em PDF ou Excel** direto aqui.\n\n"
                "⚠️ *Esta recomendação não substitui o laudo de Engenheiro Agrônomo habilitado (CREA).*"
            ),
        })

    # Upload de imagem (sidebar)
    with st.sidebar:
        st.markdown("### 📎 Enviar Laudo / Foto")
        uploaded = st.file_uploader(
            "Foto do laudo, planta ou solo",
            type=["jpg", "jpeg", "png", "webp"],
            help="Envie foto do laudo de solo, planta com deficiência ou área problema"
        )
        if uploaded:
            st.image(uploaded, caption="Imagem carregada", use_container_width=True)

        st.divider()
        st.markdown("### 💡 Exemplos de perguntas")
        exemplos = [
            "Interprete este laudo de solo",
            "Calcule calagem para pH 6,0",
            "Recomende NPK para milho 120 sc/ha",
            "O que causa clorose nas folhas?",
            "Qual formulado para soja no cerrado?",
            "Gere um relatório em PDF",
        ]
        for ex in exemplos:
            if st.button(ex, use_container_width=True, key=f"ex_{ex}"):
                st.session_state.exemplo_selecionado = ex
                st.rerun()

        st.divider()
        if st.button("🗑️ Limpar conversa", use_container_width=True):
            st.session_state.messages = []
            st.rerun()

    DISCLAIMER = "⚠️ **Esta recomendação não substitui o laudo de Engenheiro Agrônomo habilitado (CREA).**"

    def exibe_msg_assistente(content: str, show_disclaimer: bool = True):
        # Remove disclaimer inline que o modelo possa ter incluído
        texto = re.sub(r'⚠️\s*\*?Esta recomendação não substitui.*?CREA\.\*?\n?', '', content, flags=re.IGNORECASE).strip()
        st.markdown(texto)
        if show_disclaimer:
            st.markdown(
                f'<div class="disclaimer-box">{DISCLAIMER}</div>',
                unsafe_allow_html=True
            )

    # Exibe histórico
    for i, msg in enumerate(st.session_state.messages):
        avatar = "🌱" if msg["role"] == "assistant" else "👤"
        with st.chat_message(msg["role"], avatar=avatar):
            if msg["role"] == "assistant":
                # Mensagem de boas-vindas (índice 0) mostra disclaimer simples
                exibe_msg_assistente(msg["content"], show_disclaimer=(i > 0))
            else:
                st.markdown(msg["content"])

    # Botões de download da última resposta do assistente
    assistant_msgs = [m for m in st.session_state.messages if m["role"] == "assistant"]
    if len(assistant_msgs) > 1:  # exclui mensagem de boas-vindas
        last_ai = assistant_msgs[-1]["content"]
        nome_data = datetime.now().strftime("%Y%m%d_%H%M")

        pdf_bytes   = markdown_to_pdf(last_ai, "Análise Agrônomo IA")
        excel_bytes = response_to_excel(last_ai, "Análise Agrônomo IA")

        link_pdf   = make_download_link(pdf_bytes,   f"agronomo_ia_{nome_data}.pdf",  "📥 Baixar PDF",   "")
        link_excel = make_download_link(excel_bytes, f"agronomo_ia_{nome_data}.xlsx", "📊 Baixar Excel", "green")

        st.markdown(
            f'<div style="margin-bottom:8px">{link_pdf} {link_excel}</div>',
            unsafe_allow_html=True
        )

    # Input
    image_b64, image_mime = None, None

    # Exemplo selecionado na sidebar
    prompt = st.session_state.pop("exemplo_selecionado", None)

    # Chat input principal
    chat_input = st.chat_input(
        "Envie o laudo, faça uma pergunta ou peça um relatório...",
        key="chat_input"
    )
    if chat_input:
        prompt = chat_input

    if prompt:
        # Prepara imagem se houver
        if uploaded:
            img_bytes = uploaded.read()
            image_b64  = base64.b64encode(img_bytes).decode()
            image_mime = uploaded.type
            display_content = f"📷 *[Imagem enviada: {uploaded.name}]*\n\n{prompt}"
        else:
            display_content = prompt

        # Mostra mensagem do usuário
        with st.chat_message("user", avatar="👤"):
            st.markdown(display_content)

        st.session_state.messages.append({"role": "user", "content": display_content})

        # Chama Claude
        resposta = call_claude(st.session_state.messages, image_b64, image_mime)

        # Mostra resposta
        with st.chat_message("assistant", avatar="🌱"):
            exibe_msg_assistente(resposta)

        st.session_state.messages.append({"role": "assistant", "content": resposta})
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if not check_auth():
    st.stop()

main()
