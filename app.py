import sys
import csv
import subprocess
import tempfile
from pathlib import Path

import streamlit as st


BASE_DIR = Path(__file__).parent
SCRIPT_PATH = BASE_DIR / "generar_excel.py"

TEMPLATE_CANDIDATES = [
    BASE_DIR / "B2B TEMPLATE- GAMES INTEGRATIONS.xlsx",
    *BASE_DIR.glob("*GAMES INTEGRATIONS*.xlsx"),
]


def find_template():
    for path in TEMPLATE_CANDIDATES:
        if path.exists() and path.suffix.lower() == ".xlsx":
            return path
    return None


def ensure_xlsx_name(name):
    name = name.strip() or "B2B_completed.xlsx"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def read_audit_csv(audit_path):
    if not audit_path.exists():
        return []

    with audit_path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def mark_excel_downloaded():
    st.session_state.excel_downloaded = True


def mark_audit_downloaded():
    st.session_state.audit_downloaded = True


def reset_process():
    st.session_state.excel_bytes = None
    st.session_state.audit_bytes = None
    st.session_state.audit_rows = []
    st.session_state.process_done = False
    st.session_state.process_error = None
    st.session_state.excel_downloaded = False
    st.session_state.audit_downloaded = False
    st.session_state.output_name = "B2B_completed.xlsx"
    st.session_state.uploader_key += 1


if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

if "audit_bytes" not in st.session_state:
    st.session_state.audit_bytes = None

if "audit_rows" not in st.session_state:
    st.session_state.audit_rows = []

if "process_done" not in st.session_state:
    st.session_state.process_done = False

if "process_error" not in st.session_state:
    st.session_state.process_error = None

if "excel_downloaded" not in st.session_state:
    st.session_state.excel_downloaded = False

if "audit_downloaded" not in st.session_state:
    st.session_state.audit_downloaded = False

if "output_name" not in st.session_state:
    st.session_state.output_name = "B2B_completed.xlsx"


st.set_page_config(
    page_title="Certificador de Juegos Perú",
    layout="centered"
)

st.title("Certificador de Juegos - Perú")
st.write(
    "Carga uno o varios certificados PDF y genera automáticamente "
    "el Excel B2B completado para certificación de juegos en Perú."
)

template_path = find_template()

with st.expander("Verificación de archivos del sistema"):
    st.write(f"Carpeta base: `{BASE_DIR}`")
    st.write(f"Script encontrado: `{SCRIPT_PATH.exists()}`")
    st.write(f"Ruta del script: `{SCRIPT_PATH}`")
    st.write(f"Plantilla encontrada: `{template_path is not None}`")
    st.write(f"Ruta de plantilla: `{template_path}`")

uploaded_files = st.file_uploader(
    "Sube los certificados PDF",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_uploader_{st.session_state.uploader_key}",
)

output_name_input = st.text_input(
    "Nombre del Excel de salida",
    value=st.session_state.output_name,
)

generate = st.button("Generar Excel", type="primary")

if generate:
    output_name = ensure_xlsx_name(output_name_input)

    st.session_state.excel_bytes = None
    st.session_state.audit_bytes = None
    st.session_state.audit_rows = []
    st.session_state.process_done = False
    st.session_state.process_error = None
    st.session_state.excel_downloaded = False
    st.session_state.audit_downloaded = False
    st.session_state.output_name = output_name

    if not uploaded_files:
        st.session_state.process_error = "Debes cargar al menos un PDF."
    elif not SCRIPT_PATH.exists():
        st.session_state.process_error = f"No encontré el script: {SCRIPT_PATH}"
    elif template_path is None:
        st.session_state.process_error = "No encontré la plantilla Excel B2B en la carpeta del proyecto."
    else:
        with st.spinner("Procesando certificados..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir)

                pdf_dir = tmp_path / "PDFS"
                pdf_dir.mkdir(parents=True, exist_ok=True)

                output_path = tmp_path / output_name
                audit_path = tmp_path / "audit_certificacion.csv"

                for uploaded_file in uploaded_files:
                    pdf_path = pdf_dir / uploaded_file.name
                    pdf_path.write_bytes(uploaded_file.read())

                command = [
                    sys.executable,
                    str(SCRIPT_PATH),
                    "--template",
                    str(template_path),
                    "--pdf-dir",
                    str(pdf_dir),
                    "--output",
                    str(output_path),
                    "--audit",
                    str(audit_path),
                ]

                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                )

                if result.returncode != 0:
                    st.session_state.process_error = result.stderr or result.stdout
                elif not output_path.exists():
                    st.session_state.process_error = "El proceso terminó, pero no se generó el archivo Excel."
                else:
                    st.session_state.excel_bytes = output_path.read_bytes()
                    st.session_state.process_done = True

                    if audit_path.exists():
                        st.session_state.audit_bytes = audit_path.read_bytes()
                        st.session_state.audit_rows = read_audit_csv(audit_path)

                    # Limpia visualmente los PDFs cargados después de procesar.
                    st.session_state.uploader_key += 1
                    st.rerun()


if st.session_state.process_error:
    st.error("Hubo un error procesando los certificados.")
    st.code(st.session_state.process_error)

if st.session_state.process_done:
    audit_rows = st.session_state.audit_rows
    failed_rows = [
        row for row in audit_rows
        if row.get("status", "").upper() not in ("OK", "")
    ]

    if failed_rows:
        st.error("El proceso terminó, pero algunos PDFs requieren revisión.")

        for row in failed_rows:
            pdf_name = row.get("pdf", "PDF desconocido")
            status = row.get("status", "REVISAR")
            message = row.get("message", "")

            st.write(f"**{pdf_name}** — Estado: `{status}`")
            if message:
                st.caption(message)

        with st.expander("Ver auditoría completa"):
            st.dataframe(audit_rows, use_container_width=True)

    else:
        st.success("Todo procesado exitosamente.")

        if audit_rows:
            total_pdfs = len(audit_rows)
            total_games = sum(
                int(row.get("extracted_games") or 0)
                for row in audit_rows
            )

            st.info(
                f"Se procesaron {total_pdfs} PDF(s) y se extrajeron "
                f"{total_games} juego(s)."
            )

            with st.expander("Ver auditoría del proceso"):
                st.dataframe(audit_rows, use_container_width=True)

    col1, col2 = st.columns(2)

    with col1:
        if st.session_state.excel_bytes:
            st.download_button(
                label=(
                    "Excel descargado"
                    if st.session_state.excel_downloaded
                    else "Descargar Excel completado"
                ),
                data=st.session_state.excel_bytes,
                file_name=st.session_state.output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=st.session_state.excel_downloaded,
                on_click=mark_excel_downloaded,
            )

    with col2:
        if st.session_state.audit_bytes:
            st.download_button(
                label=(
                    "Auditoría descargada"
                    if st.session_state.audit_downloaded
                    else "Descargar auditoría CSV"
                ),
                data=st.session_state.audit_bytes,
                file_name="audit_certificacion.csv",
                mime="text/csv",
                disabled=st.session_state.audit_downloaded,
                on_click=mark_audit_downloaded,
            )

    st.divider()

    if st.button("Nuevo procesamiento"):
        reset_process()
        st.rerun()