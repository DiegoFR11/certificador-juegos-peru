from pathlib import Path
import shutil
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

from generar_excel import process


APP_TITLE = "Certificador de Juegos - Perú"
APP_SUBTITLE = (
    "Carga certificados PDF, extrae la información de juegos, completa la plantilla B2B "
    "y genera una auditoría de validación."
)

TEMPLATE_PATH = Path("B2B TEMPLATE- GAMES INTEGRATIONS.xlsx")
LOGO_PATH = Path("assets/logo_micasino.png")

# Paleta MiCasino basada en amarillo, blanco y negro.
YELLOW = "#FFC72C"
YELLOW_DARK = "#DDA600"
BLACK = "#111111"
WHITE = "#FFFFFF"
LIGHT_BG = "#FFF9E8"
BORDER = "#E5E7EB"


AUDIT_COLUMNS = [
    "pdf",
    "document_type",
    "report_reference",
    "expected_games",
    "extracted_games",
    "status",
    "message",
]


DOC_TYPE_LABELS = {
    "QUINEL_GAME_CERTIFICATE": "Certificado de juego - QUINEL Ltd",
    "GLI_GAME_CERTIFICATE": "Certificado de juego - Gaming Laboratories International (GLI)",
    "RNG_GNA": "Certificado RNG/GNA",
    "MINCETUR_RESOLUTION": "Resolución Directoral MINCETUR",
    "UNKNOWN": "No identificado",
    "ERROR": "Error de procesamiento",
}


def init_session_state():
    defaults = {
        "processed": False,
        "excel_bytes": None,
        "audit_bytes": None,
        "audit_df": pd.DataFrame(columns=AUDIT_COLUMNS),
        "audit_rows": [],
        "excel_downloaded": False,
        "audit_downloaded": False,
        "uploader_key": 0,
        "last_error": None,
        "processed_file_names": [],
        "processed_at": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_results():
    st.session_state.processed = False
    st.session_state.excel_bytes = None
    st.session_state.audit_bytes = None
    st.session_state.audit_df = pd.DataFrame(columns=AUDIT_COLUMNS)
    st.session_state.audit_rows = []
    st.session_state.excel_downloaded = False
    st.session_state.audit_downloaded = False
    st.session_state.last_error = None
    st.session_state.processed_file_names = []
    st.session_state.processed_at = None


def reset_uploader():
    st.session_state.uploader_key += 1


def mark_excel_downloaded():
    st.session_state.excel_downloaded = True


def mark_audit_downloaded():
    st.session_state.audit_downloaded = True


def load_file_bytes(path):
    with open(path, "rb") as file:
        return file.read()


def audit_rows_to_dataframe(audit_rows):
    if not audit_rows:
        return pd.DataFrame(columns=AUDIT_COLUMNS)

    df = pd.DataFrame(audit_rows)

    for column in AUDIT_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    return df[AUDIT_COLUMNS]


def get_result_summary(df):
    if df is None or df.empty:
        return {
            "total_pdfs": 0,
            "ok_count": 0,
            "revisar_count": 0,
            "error_count": 0,
            "total_games": 0,
            "overall_status": "SIN_RESULTADOS",
        }

    total_pdfs = len(df)
    ok_count = int((df["status"] == "OK").sum()) if "status" in df.columns else 0
    revisar_count = int((df["status"] == "REVISAR").sum()) if "status" in df.columns else 0
    error_count = int((df["status"] == "ERROR").sum()) if "status" in df.columns else 0
    total_games = int(pd.to_numeric(df.get("extracted_games", 0), errors="coerce").fillna(0).sum())

    if error_count > 0:
        overall_status = "ERROR"
    elif revisar_count > 0:
        overall_status = "REVISAR"
    else:
        overall_status = "OK"

    return {
        "total_pdfs": total_pdfs,
        "ok_count": ok_count,
        "revisar_count": revisar_count,
        "error_count": error_count,
        "total_games": total_games,
        "overall_status": overall_status,
    }


def render_css():
    st.markdown(
        f"""
        <style>
            :root {{
                --micasino-yellow: {YELLOW};
                --micasino-yellow-dark: {YELLOW_DARK};
                --micasino-black: {BLACK};
                --micasino-white: {WHITE};
                --micasino-light-bg: {LIGHT_BG};
                --micasino-border: {BORDER};
            }}

            .stApp {{
                background: linear-gradient(180deg, #FFFFFF 0%, #FFF9E8 100%);
            }}

            section[data-testid="stSidebar"] {{
                background: #111111;
                border-right: 4px solid var(--micasino-yellow);
            }}

            section[data-testid="stSidebar"] * {{
                color: #FFFFFF !important;
            }}

            section[data-testid="stSidebar"] .stAlert * {{
                color: inherit !important;
            }}

            section[data-testid="stSidebar"] hr {{
                border-color: rgba(255, 199, 44, 0.35);
            }}

            /* Botones del sidebar: corrige el problema de texto blanco sobre fondo claro. */
            section[data-testid="stSidebar"] div.stButton > button {{
                background: var(--micasino-yellow) !important;
                color: #111111 !important;
                border: 1px solid var(--micasino-yellow) !important;
                border-radius: 14px !important;
                height: 3rem !important;
                font-weight: 900 !important;
                box-shadow: 0 7px 16px rgba(255, 199, 44, 0.22) !important;
            }}

            section[data-testid="stSidebar"] div.stButton > button:hover {{
                background: var(--micasino-yellow-dark) !important;
                color: #111111 !important;
                border: 1px solid #FFFFFF !important;
            }}

            section[data-testid="stSidebar"] div.stButton > button:disabled,
            section[data-testid="stSidebar"] div.stButton > button[disabled] {{
                background: #3A3A3A !important;
                color: #BDBDBD !important;
                border: 1px solid #555555 !important;
                opacity: 1 !important;
            }}

            .brand-hero {{
                background: #111111;
                border: 2px solid var(--micasino-yellow);
                border-radius: 24px;
                padding: 1.4rem 1.6rem;
                margin-bottom: 1.4rem;
                box-shadow: 0 16px 36px rgba(17, 17, 17, 0.18);
            }}

            .brand-kicker {{
                color: var(--micasino-yellow);
                font-size: 0.78rem;
                font-weight: 800;
                letter-spacing: 0.12em;
                text-transform: uppercase;
                margin-bottom: 0.2rem;
            }}

            .brand-title {{
                color: #FFFFFF;
                font-size: 2.25rem;
                font-weight: 900;
                line-height: 1.08;
                margin: 0;
            }}

            .brand-subtitle {{
                color: #F3F4F6;
                font-size: 1rem;
                margin-top: 0.7rem;
                max-width: 820px;
            }}

            .brand-pill {{
                display: inline-flex;
                align-items: center;
                gap: 0.4rem;
                background: var(--micasino-yellow);
                color: #111111;
                font-weight: 800;
                border-radius: 999px;
                padding: 0.42rem 0.8rem;
                font-size: 0.82rem;
                margin-top: 0.9rem;
            }}

            .section-card {{
                background: #FFFFFF;
                border: 1px solid var(--micasino-border);
                border-radius: 20px;
                padding: 1.2rem 1.3rem;
                box-shadow: 0 10px 24px rgba(17, 17, 17, 0.06);
                margin-bottom: 1rem;
            }}

            .info-card {{
                padding: 1rem 1.1rem;
                border-radius: 16px;
                background: #FFFFFF;
                border: 1px solid var(--micasino-border);
                border-left: 6px solid var(--micasino-yellow);
                margin-bottom: 1rem;
            }}

            .success-card {{
                padding: 1rem 1.1rem;
                border-radius: 16px;
                background: #E8F7EF;
                border: 1px solid #B7E4C7;
                color: #146C43;
                margin-bottom: 1rem;
                font-weight: 650;
            }}

            .warning-card {{
                padding: 1rem 1.1rem;
                border-radius: 16px;
                background: #FFF4DB;
                border: 1px solid #FFD27D;
                color: #8A5A00;
                margin-bottom: 1rem;
                font-weight: 650;
            }}

            .error-card {{
                padding: 1rem 1.1rem;
                border-radius: 16px;
                background: #FDECEC;
                border: 1px solid #F5B5B5;
                color: #B42318;
                margin-bottom: 1rem;
                font-weight: 650;
            }}

            .next-step-card {{
                padding: 1.1rem 1.2rem;
                border-radius: 18px;
                background: #111111;
                color: #FFFFFF;
                border: 2px solid var(--micasino-yellow);
                box-shadow: 0 10px 24px rgba(17, 17, 17, 0.12);
                margin-top: 1rem;
                margin-bottom: 1rem;
            }}

            .next-step-card strong {{
                color: var(--micasino-yellow);
            }}

            .file-chip {{
                background: #111111;
                color: #FFFFFF;
                border-radius: 999px;
                padding: 0.42rem 0.75rem;
                display: inline-block;
                margin: 0.25rem 0.25rem 0.25rem 0;
                font-size: 0.86rem;
                border: 1px solid var(--micasino-yellow);
            }}

            .step-flow {{
                display: flex;
                gap: 0.6rem;
                flex-wrap: wrap;
                margin-bottom: 1rem;
            }}

            .step-item {{
                background: #FFFFFF;
                border: 1px solid var(--micasino-border);
                border-radius: 999px;
                padding: 0.5rem 0.9rem;
                color: #111111;
                font-weight: 800;
                font-size: 0.88rem;
            }}

            .step-item.active {{
                background: var(--micasino-yellow);
                border: 1px solid #111111;
            }}

            div[data-testid="stMetric"] {{
                background: #FFFFFF;
                border: 1px solid var(--micasino-border);
                border-bottom: 4px solid var(--micasino-yellow);
                padding: 1rem;
                border-radius: 18px;
                box-shadow: 0 8px 18px rgba(17, 17, 17, 0.05);
            }}

            div[data-testid="stMetric"] label {{
                color: #111111 !important;
                font-weight: 800 !important;
            }}

            div.stButton > button,
            div.stDownloadButton > button {{
                border-radius: 14px;
                height: 3rem;
                font-weight: 800;
                border: 1px solid #111111;
                box-shadow: 0 7px 16px rgba(17, 17, 17, 0.12);
            }}

            div.stButton > button[kind="primary"],
            div.stDownloadButton > button[kind="primary"] {{
                background: var(--micasino-yellow);
                color: #111111;
                border: 1px solid #111111;
            }}

            div.stButton > button[kind="primary"]:hover,
            div.stDownloadButton > button[kind="primary"]:hover {{
                background: var(--micasino-yellow-dark);
                color: #111111;
                border: 1px solid #111111;
            }}

            div.stButton > button:disabled,
            div.stDownloadButton > button:disabled {{
                background: #E5E7EB !important;
                color: #4B5563 !important;
                border: 1px solid #9CA3AF !important;
                opacity: 1 !important;
                box-shadow: none !important;
            }}

            .stTabs [data-baseweb="tab-list"] {{
                gap: 0.5rem;
            }}

            .stTabs [data-baseweb="tab"] {{
                background: #FFFFFF;
                border: 1px solid var(--micasino-border);
                border-radius: 999px;
                padding: 0.6rem 1rem;
                color: #111111;
                font-weight: 800;
            }}

            .stTabs [aria-selected="true"] {{
                background: var(--micasino-yellow) !important;
                color: #111111 !important;
                border: 1px solid #111111 !important;
            }}

            .small-muted {{
                font-size: 0.85rem;
                color: #6B7280;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_logo(width=76):
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=width)
    else:
        st.markdown(
            """
            <div style="
                width:76px;height:76px;border-radius:22px;background:#FFC72C;
                color:#111111;display:flex;align-items:center;justify-content:center;
                font-weight:900;font-size:1.8rem;border:2px solid #FFFFFF;">
                M
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_header():
    col_logo, col_text = st.columns([0.11, 0.89], vertical_alignment="center")

    with col_logo:
        render_logo(width=82)

    with col_text:
        st.markdown(
            f"""
            <div class="brand-hero">
                <div class="brand-kicker">MiCasino.com · Automatización de certificaciones</div>
                <h1 class="brand-title">{APP_TITLE}</h1>
                <div class="brand-subtitle">{APP_SUBTITLE}</div>
                <div class="brand-pill">🎰 Perú · B2B Games Integrations</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_sidebar():
    with st.sidebar:
        render_logo(width=92)
        st.markdown("## MiCasino.com")
        st.caption("Certificador de Juegos - Perú")

        st.divider()

        st.markdown("### Plantilla activa")
        if TEMPLATE_PATH.exists():
            st.success(TEMPLATE_PATH.name)
        else:
            st.error("No se encontró la plantilla Excel.")

        st.divider()

        st.markdown("### Entidades certificadoras soportadas")
        st.caption("✅ QUINEL Ltd")
        st.caption("✅ Gaming Laboratories International (GLI)")

        st.divider()

        st.markdown("### Acciones")
        if st.button("Limpiar resultados", use_container_width=True):
            reset_results()
            reset_uploader()
            st.rerun()


def style_audit_table(df):
    if df.empty or "status" not in df.columns:
        return df

    def color_status(row):
        status = str(row.get("status", "")).upper()

        if status == "OK":
            return ["background-color: #E8F7EF; color: #146C43; font-weight: 650"] * len(row)

        if status == "REVISAR":
            return ["background-color: #FFF4DB; color: #8A5A00; font-weight: 650"] * len(row)

        if status == "ERROR":
            return ["background-color: #FDECEC; color: #B42318; font-weight: 650"] * len(row)

        return [""] * len(row)

    return df.style.apply(color_status, axis=1)


def render_step_flow(active_step):
    steps = [
        (1, "Cargar PDFs"),
        (2, "Procesar"),
        (3, "Validar auditoría"),
        (4, "Descargar archivos"),
    ]
    html = '<div class="step-flow">'
    for number, label in steps:
        cls = "step-item active" if number == active_step else "step-item"
        html += f'<span class="{cls}">{number}. {label}</span>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def render_metrics(summary):
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("PDFs", summary["total_pdfs"])
    col2.metric("OK", summary["ok_count"])
    col3.metric("Revisar", summary["revisar_count"])
    col4.metric("Errores", summary["error_count"])
    col5.metric("Juegos", summary["total_games"])


def render_status_message(summary):
    if summary["overall_status"] == "ERROR":
        st.markdown(
            """
            <div class="error-card">
                Se generó la salida, pero hay certificados con error. Revisa la auditoría antes de usar el Excel.
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif summary["overall_status"] == "REVISAR":
        st.markdown(
            """
            <div class="warning-card">
                El Excel fue generado, pero hay certificados marcados como REVISAR. Valida la auditoría antes de continuar.
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif summary["overall_status"] == "OK":
        st.markdown(
            """
            <div class="success-card">
                Proceso completado correctamente. Ya puedes descargar el Excel completado y el CSV de auditoría.
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_download_buttons(prefix=""):
    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            label="Descargar Excel completado",
            data=st.session_state.excel_bytes,
            file_name="B2B_TEMPLATE_GAMES_INTEGRATIONS_COMPLETADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            disabled=st.session_state.excel_downloaded,
            on_click=mark_excel_downloaded,
            key=f"{prefix}_download_excel",
        )

        if st.session_state.excel_downloaded:
            st.caption("Excel descargado. Botón bloqueado para evitar descargas duplicadas.")

    with col2:
        st.download_button(
            label="Descargar CSV de auditoría",
            data=st.session_state.audit_bytes,
            file_name="auditoria_certificados.csv",
            mime="text/csv",
            type="secondary",
            use_container_width=True,
            disabled=st.session_state.audit_downloaded,
            on_click=mark_audit_downloaded,
            key=f"{prefix}_download_audit",
        )

        if st.session_state.audit_downloaded:
            st.caption("Auditoría descargada. Botón bloqueado para evitar descargas duplicadas.")


def render_processed_summary_on_upload_tab():
    df = st.session_state.audit_df.copy()
    summary = get_result_summary(df)

    st.markdown(
        """
        <div class="next-step-card">
            <strong>Resultado listo.</strong><br>
            La carga ya fue procesada. Revisa el resumen, descarga los archivos o abre la pestaña de auditoría para validar el detalle por PDF.
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.session_state.processed_at:
        st.caption(f"Último procesamiento: {st.session_state.processed_at}")

    if st.session_state.processed_file_names:
        with st.expander("PDFs procesados", expanded=False):
            for file_name in st.session_state.processed_file_names:
                st.write(f"📄 {file_name}")

    render_metrics(summary)
    st.write("")
    render_status_message(summary)
    render_download_buttons(prefix="upload_tab")

    with st.expander("Vista rápida de auditoría", expanded=summary["overall_status"] != "OK"):
        st.dataframe(style_audit_table(df), use_container_width=True, hide_index=True)

    st.info("Para iniciar otro lote de PDFs, usa el botón 'Limpiar resultados' en el panel lateral.")


def render_upload_tab():
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    render_step_flow(active_step=1)
    st.subheader("1. Carga de certificados PDF")
    st.write("Sube uno o varios certificados. La herramienta procesará todos en un único Excel.")

    if st.session_state.processed:
        render_processed_summary_on_upload_tab()
        st.markdown("</div>", unsafe_allow_html=True)
        return

    uploaded_files = st.file_uploader(
        "Arrastra o selecciona tus certificados PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.uploader_key}",
    )

    if not uploaded_files:
        st.markdown(
            """
            <div class="info-card">
                Carga certificados PDF para iniciar el proceso. Cuando los selecciones, aparecerá el botón
                <strong>Procesar certificados</strong> en esta misma pantalla.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
        return

    st.markdown(f"**PDFs cargados:** {len(uploaded_files)}")

    chips = "".join(f'<span class="file-chip">📄 {file.name}</span>' for file in uploaded_files)
    st.markdown(chips, unsafe_allow_html=True)

    st.markdown(
        """
        <div class="next-step-card">
            <strong>Siguiente paso:</strong> haz clic en <strong>Procesar certificados</strong>.
            Al finalizar, verás el resumen, la auditoría y los botones de descarga en esta misma pantalla.
        </div>
        """,
        unsafe_allow_html=True,
    )

    process_clicked = st.button(
        "Procesar certificados",
        type="primary",
        use_container_width=True,
        disabled=not TEMPLATE_PATH.exists(),
    )

    st.markdown("</div>", unsafe_allow_html=True)

    if process_clicked:
        reset_results()

        with st.spinner("Procesando certificados..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmpdir = Path(tmpdir)

                    pdf_dir = tmpdir / "pdfs"
                    pdf_dir.mkdir(parents=True, exist_ok=True)

                    pdf_paths = []
                    processed_file_names = []

                    for uploaded_file in uploaded_files:
                        pdf_path = pdf_dir / uploaded_file.name
                        with open(pdf_path, "wb") as file:
                            file.write(uploaded_file.getbuffer())
                        pdf_paths.append(pdf_path)
                        processed_file_names.append(uploaded_file.name)

                    template_copy = tmpdir / TEMPLATE_PATH.name
                    shutil.copy(TEMPLATE_PATH, template_copy)

                    output_excel = tmpdir / "B2B_TEMPLATE_GAMES_INTEGRATIONS_COMPLETADO.xlsx"
                    output_audit = tmpdir / "auditoria_certificados.csv"

                    audit_rows = process(
                        template=template_copy,
                        pdfs=pdf_paths,
                        output=output_excel,
                        audit=output_audit,
                        strict=False,
                    )

                    st.session_state.excel_bytes = load_file_bytes(output_excel)
                    st.session_state.audit_bytes = load_file_bytes(output_audit)
                    st.session_state.audit_rows = audit_rows
                    st.session_state.audit_df = audit_rows_to_dataframe(audit_rows)
                    st.session_state.processed = True
                    st.session_state.processed_file_names = processed_file_names
                    st.session_state.processed_at = datetime.now().strftime("%d-%m-%Y %H:%M")

                    reset_uploader()

                st.rerun()

            except Exception as exc:
                st.session_state.last_error = str(exc)
                st.error(f"Ocurrió un error procesando los certificados: {exc}")


def render_results_tab():
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    render_step_flow(active_step=3)
    st.subheader("2. Auditoría de procesamiento")

    if st.session_state.last_error:
        st.markdown(
            f"""
            <div class="error-card">
                <strong>Error:</strong> {st.session_state.last_error}
            </div>
            """,
            unsafe_allow_html=True,
        )

    if not st.session_state.processed:
        st.markdown(
            """
            <div class="info-card">
                Aún no hay resultados. Primero carga y procesa uno o varios certificados PDF.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
        return

    df = st.session_state.audit_df.copy()
    summary = get_result_summary(df)

    render_metrics(summary)
    st.write("")
    render_status_message(summary)

    display_df = df.copy()
    if "document_type" in display_df.columns:
        display_df["document_type"] = display_df["document_type"].map(DOC_TYPE_LABELS).fillna(display_df["document_type"])

    st.dataframe(
        style_audit_table(display_df),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("</div>", unsafe_allow_html=True)


def render_downloads_tab():
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    render_step_flow(active_step=4)
    st.subheader("3. Descargas")

    if not st.session_state.processed:
        st.markdown(
            """
            <div class="info-card">
                Las descargas estarán disponibles después de procesar los certificados.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
        return

    df = st.session_state.audit_df.copy()
    summary = get_result_summary(df)
    render_status_message(summary)
    render_download_buttons(prefix="downloads_tab")

    st.markdown("</div>", unsafe_allow_html=True)


def main():
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="🎰",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    init_session_state()
    render_css()
    render_sidebar()
    render_header()

    tab_upload, tab_results, tab_downloads = st.tabs(
        [
            "📤 Cargar y procesar",
            "📋 Auditoría",
            "⬇️ Descargas",
        ]
    )

    with tab_upload:
        render_upload_tab()

    with tab_results:
        render_results_tab()

    with tab_downloads:
        render_downloads_tab()


if __name__ == "__main__":
    main()
