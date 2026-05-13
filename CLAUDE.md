# Certificador de Juegos — Perú · Contexto del Proyecto

## Instrucción para Claude

Eres un asistente experto trabajando en el proyecto **Certificador de Juegos — Perú** de MiCasino.com.
Conoces el código completo de la app. Cuando el usuario te pida cambios, mejoras o nuevas funcionalidades:

1. Responde siempre en **español**.
2. Entrega el **archivo completo modificado**, listo para reemplazar directamente en el repositorio. Nunca entregues diffs parciales.
3. Si el cambio afecta a más de un archivo, entrega todos los archivos afectados completos.
4. Antes de modificar, explica brevemente **qué vas a cambiar y por qué**.
5. Si detectas efectos secundarios o riesgos en otros módulos, avísalos.
6. Respeta el estilo de código existente: funciones pequeñas, docstrings cortos, regex nombrados con RE_, sesión state centralizado en `init_session_state()`.

---

## Descripción General

Herramienta web interna de **MiCasino.com** construida con **Streamlit** que automatiza la gestión de documentación regulatoria de juegos de casino en Perú.

**Problema que resuelve:** los equipos de cumplimiento recibían certificados PDF de entidades internacionales (GLI, QUINEL) y Resoluciones Directorales del MINCETUR, y debían copiar manualmente los datos a plantillas Excel regulatorias — proceso lento y propenso a errores.

**Solución:** extrae automáticamente la información de los PDFs con PyMuPDF + regex, la mapea a la plantilla B2B y genera un Excel listo más un CSV de auditoría.

**Repositorio:** https://github.com/DiegoFR11/certificador-juegos-peru  
**Stack:** Python · Streamlit · PyMuPDF · openpyxl · pandas

---

## Archivos del Proyecto

| Archivo | Rol |
|---|---|
| `app.py` | UI Streamlit + lógica completa del módulo MINCETUR |
| `generar_excel.py` | Motor de extracción PDF → Excel (certificados GLI/QUINEL) |
| `B2B TEMPLATE- GAMES INTEGRATIONS.xlsx` | Plantilla regulatoria base |
| `assets/logo_micasino.png` | Logo de MiCasino |
| `.streamlit/config.toml` | Configuración de tema visual |
| `requirements.txt` | Dependencias Python |

---

## Dependencias

```
PyMuPDF==1.24.14
openpyxl==3.1.5
pandas
streamlit
```

---

## Branding y Estilos

| Variable | Valor |
|---|---|
| Amarillo principal | `#FFC629` |
| Negro | `#0B0B0B` |
| Fondo principal | `#FFF9EA` |
| Fondo suave | `#FFF3C4` |
| Texto | `#111827` |

- Sidebar oscuro (`#0B0B0B`) con texto blanco.
- Área principal en modo claro forzado via CSS (`color-scheme: light`).
- Header renderizado con `components.html` para aislar CSS del modo oscuro del navegador.
- Botones: amarillo con texto negro, `border-radius: 14px`, `font-weight: 800`.
- Tabs activos: fondo amarillo, texto negro bold.

---

## Flujo 1 — Tab "📥 Certificados" (GLI / QUINEL)

```
Usuario sube PDFs
       ↓
generar_excel.py · process()
  ├── read_pdf_text()          → extrae texto plano con PyMuPDF
  ├── detect_document_type()   → identifica GLI / QUINEL / RNG / MINCETUR
  ├── extract_header()         → proveedor, fabricante, referencia, fechas, RNG
  ├── extract_games()          → lista de juegos: nombre, versión, código único
  ├── fill_excel()             → escribe sobre copia de la plantilla B2B
  │     └── detecta columnas editables marcadas "to be completed by the client"
  └── write_audit_csv()        → CSV con estado por PDF (OK / REVISAR / ERROR)
       ↓
Descarga: Excel B2B + CSV auditoría (nombres con timestamp)
```

---

## Flujo 2 — Tab "🏛️ Resoluciones MINCETUR"

```
Usuario sube PDFs (Resoluciones Directorales)
       ↓
app.py · process_mincetur_resolutions()
  ├── read_pdf_text()                    → extrae texto
  ├── detect_tipo_componente()           → "Programa de juego" / "Plataforma tecnológica"
  │                                         "Sistema progresivo" / "Casino en vivo"
  ├── extract_resolution_manufacturer()  → proveedor principal del documento
  ├── extract_mincetur_resolution_rows() → tabla Artículo 1:
  │     ├── N° registro: PJ + 7 dígitos
  │     ├── Nombre comercial del juego
  │     ├── Versión
  │     └── Código de identificación del fabricante
  ├── strip_known_manufacturer()         → limpia nombre comercial
  ├── normalize_broken_code()            → une códigos partidos entre líneas
  └── write_mincetur_excel()             → Excel regulatorio con 5 columnas
       ↓
Descarga: resoluciones_mincetur.xlsx
```

---

## Entidades Certificadoras Soportadas

| Entidad | Identificación en PDF |
|---|---|
| **GLI** (Gaming Laboratories International) | `"gaming laboratories international"`, `"gaminglabs.com"`, `"gli®"` |
| **QUINEL Ltd** | `"quinel"` |

---

## FIELD_MAP (generar_excel.py)

Mapeo entre encabezados de la plantilla B2B y campos internos:

| Encabezado plantilla | Campo interno |
|---|---|
| Game Provider | `provider` |
| Game Manufacturer | `manufacturer` |
| Game Name | `game_name` |
| Game Type | `game_type` |
| Report Reference | `report_reference` |
| Report Date | `report_date` |
| Issued by | `issued_by` |
| RNG report reference | `rng_report_reference` |
| RNG report date | `rng_report_date` |
| RNG Issued by | `rng_issued_by` |
| Sample | `sample` |
| General Result is PASS | `general_result_is_pass` |
| Unique Code | `unique_code` |

**ALWAYS_BLANK** (se dejan vacíos siempre):
- `"Report date is valid?"`
- `"Accreditation Mark and Number"`
- `"Sample"`

---

## Patrones Regex Críticos

```python
RE_GLI_REPORT_FULL  = r"\b[A-Z]{2}-\d{3}-[A-Z]{3}-\d{2}-\d{2,3}-\d{3}(?:\(\d+\))?\b"
# Ejemplo: MO-246-PPL-25-154-684

RE_GLI_REPORT_SHORT = r"\b[A-Z]{2}-\d{3}-[A-Z]{3}-\d{2}-\d{2,3}\b"
# Ejemplo: MO-246-PPL-25-154

RE_VERSION          = r"(?:cv|v)?\d+(?:\.\d+){1,3}(?:\.?r)?|N/A"
# Ejemplos: cv1.0, v2.3.1, 1.0.0.r

RE_ITEM             = r"G\d{3}"
# Ejemplo: G001

# Registro MINCETUR
r"PJ\d{7}"
# Ejemplo: PJ0012345

# Código único (hash_número)
r"[A-Za-z0-9]{12,}_[0-9.]+"
# Ejemplo: 66436d3fc069e700017e663e_961
```

---

## Session State (app.py)

### Módulo Certificados
| Key | Tipo | Descripción |
|---|---|---|
| `processed` | bool | Si se procesaron certificados |
| `excel_bytes` | bytes | Excel B2B generado |
| `audit_bytes` | bytes | CSV auditoría |
| `audit_df` | DataFrame | Auditoría como DataFrame |
| `audit_rows` | list | Filas de auditoría |
| `excel_downloaded` | bool | Bloquea re-descarga |
| `audit_downloaded` | bool | Bloquea re-descarga |
| `certificate_excel_filename` | str | Nombre con timestamp |
| `certificate_audit_filename` | str | Nombre con timestamp |
| `uploader_key` | int | Fuerza reset del file uploader |
| `last_error` | str | Último error para mostrar en UI |
| `certificate_upload_signature` | tuple | Detecta cambio de archivos cargados |

### Módulo MINCETUR
| Key | Tipo | Descripción |
|---|---|---|
| `mincetur_processed` | bool | Si se procesaron resoluciones |
| `mincetur_excel_bytes` | bytes | Excel MINCETUR generado |
| `mincetur_rows` | list | Registros extraídos |
| `mincetur_audit_rows` | list | Auditoría de PDFs |
| `mincetur_downloaded` | bool | Bloquea re-descarga |
| `mincetur_uploader_key` | int | Fuerza reset del file uploader |

---

## Convenciones del Código

- Funciones pequeñas con una sola responsabilidad.
- Docstrings cortos en español.
- Regex nombrados con prefijo `RE_` a nivel de módulo.
- Session state centralizado en `init_session_state()`.
- Uso de `tempfile.TemporaryDirectory()` para archivos temporales (nunca escribe en disco permanente).
- CSS inyectado via `st.markdown(..., unsafe_allow_html=True)` en `render_css()`.
- Botones de descarga se bloquean tras el primer clic (`on_click` + flag en session state).
