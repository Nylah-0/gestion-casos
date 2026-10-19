from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo
import hashlib
import hmac
import math
import re
import unicodedata

import pandas as pd
import streamlit as st

# El Excel se carga desde la interfaz web (barra lateral) y se guarda en el servidor,
# de modo que cualquier persona con la contraseña vea la última versión cargada.
DATA_DIR = Path(__file__).parent / "datos"
EXCEL_PATH = DATA_DIR / "INDIVIDUALES COHORTES.xlsx"
BOGOTA_TZ = ZoneInfo("America/Bogota")
REFRESH_INTERVAL = "24h"
EXPECTED_CASE_COUNT = 3746


def current_excel_file() -> Path:
    """Devuelve la ruta de la última base cargada en el servidor."""
    return EXCEL_PATH


def save_uploaded_excel(uploaded) -> tuple[bool, str]:
    """Valida el Excel cargado y lo guarda de forma atómica para todas las sesiones."""
    data = uploaded.getvalue()
    try:
        pd.ExcelFile(BytesIO(data))
    except Exception as exc:
        return False, f"El archivo no parece un Excel válido: {exc}"
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp_path = EXCEL_PATH.with_suffix(".tmp")
    tmp_path.write_bytes(data)
    tmp_path.replace(EXCEL_PATH)
    st.cache_data.clear()
    return True, ""


STATUS_COLUMN = "Estado_Registro"  # En el libro actual solo aparece "A"; no indica gestión realizada.
FOLLOWUP_COLUMN = "Estado_Gestion"  # Añádela a Excel para registrar: Pendiente, Gestionado, No contactado, etc.
FOLLOWUP_CANDIDATES = [FOLLOWUP_COLUMN, "Estado de gestión", "Estado_Gestion"]
MEDICINE_REFERRAL_CANDIDATES = ["Derivacion_Medicina", "Derivación_Medicina", "Derivacion a Medicina", "Derivación a Medicina", "Remision_Medicina", "Remisión_Medicina", "Remision a Medicina", "Remisión a Medicina", "Medicina_Remision", "Referencia_Medicina"]
MEDICINE_PENDING_CANDIDATES = ["Estado_Remision_Medicina", "Estado de remision a medicina", "Pendiente_Remision_Medicina", "Pendiente de remision a medicina"]
REFERRAL_STATUS_CANDIDATES = {
    "Medicina": [*MEDICINE_PENDING_CANDIDATES, "Estado de remision medicina", "Remitir a Medicina", "Remitir a medicina", "MEDICINA"],
    "Psicología": ["Estado_Remision_Psicologia", "Estado de remision a psicologia", "Pendiente_Remision_Psicologia", "Pendiente de remision a psicologia", "Remitir a Psicología", "Remitir a Psicologia", "PSICOLOGIA"],
    "Enfermería": ["Estado_Remision_Enfermeria", "Estado de remision a enfermeria", "Pendiente_Remision_Enfermeria", "Pendiente de remision a enfermeria", "Remitir a Enfermería", "Remitir a Enfermeria", "ENFERMERIA"],
    "Nutrición": ["Estado_Remision_Nutricion", "Estado de remision a nutricion", "Pendiente_Remision_Nutricion", "Pendiente de remision a nutricion", "Remitir a Nutrición", "Remitir a Nutricion", "NUTRICION"],
}
GENERIC_REFERRAL_STATUS_CANDIDATES = ["Estado_Remision", "Estado de remision", "Estado de referencia"]
SERVICE_COLUMN = "Serivicio_Agendado"  # Se conserva el encabezado tal como está escrito en el libro.
APPOINTMENT_DATE_COLUMN = "Fecha_de_la_Cita"
APPOINTMENT_TIME_COLUMN = "Hora_de_la_Cita"
APPOINTMENT_PLACE_COLUMN = "Lugar_de_la_Cita"
SERVICE_CANDIDATES = [SERVICE_COLUMN, "Servicio_Agendado", "Servicio agendado", "Servicio"]
 # No usar la columna general "Fecha": corresponde al registro/cohorte y no prueba que exista una cita.
APPOINTMENT_DATE_CANDIDATES = [APPOINTMENT_DATE_COLUMN, "Fecha de la cita", "Fecha cita", "Fecha_Cita"]
APPOINTMENT_TIME_CANDIDATES = [APPOINTMENT_TIME_COLUMN, "Hora de la cita", "Hora"]
ADDRESS_CANDIDATES = ["Direccion", "Dirección", "Direccion_Paciente", "Dirección_Paciente", "Direccion_Domicilio", "Dirección domicilio"]
ADDRESS_DISPLAY_CANDIDATES = [*ADDRESS_CANDIDATES, "Address"]
PCP_CANDIDATES = ["PCP", "Pcp", "PCP asignado", "Profesional de atención primaria"]
EAPB_CANDIDATES = ["Eapb", "EPS", "Asegurador", "Entidad aseguradora"]
REGIME_CANDIDATES = ["Regimen", "Régimen", "Tipo_Regimen", "Tipo de regimen"]
ALERT_COLUMNS = [f"Alerta_N°_{n}" for n in range(1, 7)]
ALERT_DESCRIPTION_COLUMNS = [f"Descripcion_Alerta_N°_{n}" for n in range(1, 7)]
PROFESSIONAL_CANDIDATES = ["Profesional_Cita", "Profesional_Agendado", "Nombre_Profesional", "Profesional"]
PROFESSIONAL_DIRECTORY = {
    "Psicología": [
        "Angélica María Forero Polanco",
        "Leidy Cuervo",
        "Liz Milena Gelvis Bonilla",
        "Gyna Paola Falla Calderón",
        "Shorlinss Freness Taylor Martin Pinzón",
    ],
    "Enfermería": ["Jenny Carolina Acosta Salcedo", "Angie Paola Romero Plata"],
    "Nutrición": ["Sergio Fernando Prieto"],
}
APPOINTMENT_SERVICES = ["Psicología", "Enfermería", "Nutrición"]
TERRITORY_DIRECTORY = {
    "SO116": "EL PORTAL DEL BRASIL",
    "SO117": "BRASILIA",
    "SO118": "BETANIA",
    "SO132": "SAN BERNARDINO II",
    "SO133": "LAS MARGARITAS",
    "SO134": "LAS MARGARITAS",
    "SO135": "SAN ANTONIO",
    "SO136": "CHICO SUR",
    "SO137": "CHICO SUR",
    "SO138": "LA INDEPENDENCIA",
    "SO139": "EL REMANSO I",
    "SO140": "VILLAS DEL PROGRESO",
    "SO141": "LAS MARGARITAS",
    "SO142": "CIUDADELA EL RECREO",
    "SO143": "CIUDADELA EL RECREO II",
    "SO144": "CIUDADELA EL RECREO II",
    "SO145": "CIUDADELA EL RECREO II",
    "SO146": "SAN BERNARDINO XIX",
    "SO147": "CIUDADELA EL RECREO",
    "SO148": "CIUDADELA EL RECREO",
    "SO149": "SAN BERNARDINO XIX",
    "SO150": "SAN BERNARDINO XXII URBANO",
    "SO151": "SAN BERNARDINO XXV URBANO",
    "SO152": "CAMPO VERDE",
    "SO153": "VILLA EMMA",
    "SO154": "LA VEGA SAN BERNARDINO",
    "SO155": "SAN BERNARDINO XVIII",
}

st.set_page_config(page_title="Gestión de casos", page_icon="📋", layout="wide")

# Contraseña sencilla en texto plano
APP_PASSWORD = "123"


def require_password() -> bool:
    if st.session_state.get("authenticated", False):
        return True

    st.title("🔐 Acceso privado")
    st.caption("Ingresa la contraseña para consultar el tablero de gestión.")
    with st.form("login_form"):
        password = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("Ingresar", use_container_width=True)

    if submitted:
        if password == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("La contraseña no es correcta.")
    return False


if not require_password():
    st.stop()

st.title("📋 Gestión de casos y agenda")
st.caption("Tablero de consulta del Excel · actualización automática una vez al día")


def norm(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip().casefold())


def find_column(columns, candidates: list[str]) -> str | None:
    def key(value: str) -> str:
        return re.sub(r"[\s_]+", " ", norm(value))

    available = {key(str(column)): str(column) for column in columns}
    return next((available[key(candidate)] for candidate in candidates if key(candidate) in available), None)


def appointment_service(value: object, professional: object = "") -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii").casefold()
    if "psicolog" in text:
        return "Psicología"
    if "enfermer" in text:
        return "Enfermería"
    if "nutri" in text:
        return "Nutrición"
    professional_key = norm(professional)
    for service, professionals in PROFESSIONAL_DIRECTORY.items():
        if any(norm(person) == professional_key for person in professionals):
            return service
    return ""


def plain_text(value: object) -> str:
    return unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii").casefold().strip()


@st.cache_data(show_spinner="Leyendo las pestañas del libro…")
def read_workbook(path_text: str, modified_at: float) -> pd.DataFrame:
    # modified_at forma parte de la clave de caché: al cargar un Excel nuevo se vuelve a leer.
    del modified_at
    workbook = pd.ExcelFile(Path(path_text))
    frames = []
    for sheet_name in workbook.sheet_names:
        frame = pd.read_excel(workbook, sheet_name=sheet_name, dtype=str).fillna("")
        # Unifica espacios repetidos en los encabezados, sin cambiar los datos.
        frame.columns = [re.sub(r"\s+", " ", str(column)).strip() for column in frame.columns]
        nonblank = frame.astype(str).apply(lambda col: col.str.strip().ne(""))
        frame = frame.loc[nonblank.any(axis=1)].copy()
        if frame.empty:
            continue
        frame["Hoja_origen"] = sheet_name
        frames.append(frame)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False).fillna("")


def parse_excel_date(value: object):
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat", "0000-00-00", "0"}:
        return pd.NaT
    # Excel puede guardar fechas como el número de días desde 30/12/1899.
    try:
        numeric = float(text)
        if numeric >= 20000:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(numeric, unit="D")
    except (ValueError, TypeError, OverflowError):
        pass
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", text):
        return pd.to_datetime(text, errors="coerce", yearfirst=True)
    slash = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if slash:
        first, second, year = map(int, slash.groups())
        dayfirst = first > 12 or (first <= 12 and second <= 12)
        if second > 12:
            dayfirst = False
        return pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def format_excel_time(value: object) -> str:
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat", "0", "0.0"}:
        return ""
    try:
        numeric = float(text)
        if 0 < numeric < 1:
            minutes = round(numeric * 24 * 60) % (24 * 60)
            return f"{minutes // 60:02d}:{minutes % 60:02d}"
    except (ValueError, TypeError):
        pass
    return text


def patient_label(row: pd.Series) -> str:
    fields = ["Primer_Nombre", "Segundo_Nombre", "Primer_Apellido", "Seundo_Apellido"]
    name = " ".join(str(row.get(col, "")).strip() for col in fields if str(row.get(col, "")).strip())
    if not name:
        name = next(
            (str(row.get(col, "")).strip() for col in ["Nombre del paciente", "Nombre_Paciente", "Paciente", "Nombre"] if str(row.get(col, "")).strip()),
            "",
        )
    code = str(row.get("Cod_Persona", "")).strip()
    return f"{name} · {code}" if name and code else name or code or "Sin identificación visible"


def get_alerts(row: pd.Series, columns: list[str]) -> str:
    found = []
    for col in columns:
        val = str(row.get(col, "")).strip()
        if val and norm(val) not in {"no", "0", "nan", "ninguna", "false"}:
            found.append(f"{col}: {val}")
    return " | ".join(found)


@st.fragment(run_every=REFRESH_INTERVAL)
def dashboard() -> None:
    # --- Carga del Excel desde la interfaz web ---
    with st.sidebar:
        st.header("Gestión")
        uploaded = st.file_uploader(
            "Cargar INDIVIDUALES COHORTES.xlsx",
            type=["xlsx"],
            key="excel_upload",
            help="Al cargar un archivo nuevo, reemplaza la base para todos los usuarios del tablero.",
        )
        if uploaded is not None:
            upload_id = f"{uploaded.name}:{uploaded.size}:{getattr(uploaded, 'file_id', '')}"
            if st.session_state.get("last_upload_id") != upload_id:
                ok, message = save_uploaded_excel(uploaded)
                if ok:
                    st.session_state["last_upload_id"] = upload_id
                    st.success("Excel cargado y base actualizada.")
                else:
                    st.error(message)

    excel_file = current_excel_file()
    try:
        if not excel_file.exists():
            st.info("Todavía no hay ninguna base cargada. Usa el cargador de la barra lateral para subir INDIVIDUALES COHORTES.xlsx.")
            return
        modified_at = excel_file.stat().st_mtime
        df = read_workbook(str(excel_file), modified_at)
        loaded_at = datetime.fromtimestamp(modified_at, tz=BOGOTA_TZ).strftime("%d/%m/%Y %I:%M %p")
        source_label = f"Excel cargado el {loaded_at}"
    except PermissionError:
        st.warning("No se pudo leer el Excel. Vuelve a cargarlo desde la barra lateral.")
        return
    except Exception as exc:
        st.error(f"No pude abrir el Excel: {exc}")
        return
    if df.empty:
        st.warning("No hay registros para mostrar.")
        return

    alert_cols = [c for c in ALERT_COLUMNS + ALERT_DESCRIPTION_COLUMNS if c in df.columns]
    df = df.copy()
    service_col = find_column(df.columns, SERVICE_CANDIDATES)
    if service_col and service_col != SERVICE_COLUMN:
        df[SERVICE_COLUMN] = df[service_col]
    elif SERVICE_COLUMN not in df.columns:
        df[SERVICE_COLUMN] = ""
    df["Paciente"] = df.apply(patient_label, axis=1)
    df["Detalle de alertas"] = df.apply(lambda row: get_alerts(row, alert_cols), axis=1)
    professional_col = find_column(df.columns, PROFESSIONAL_CANDIDATES)
    df["Profesional"] = df[professional_col].astype(str).str.strip() if professional_col else ""
    df["Servicio de agenda"] = [appointment_service(service, professional) for service, professional in zip(df[SERVICE_COLUMN], df["Profesional"])]
    referral_col = find_column(df.columns, MEDICINE_REFERRAL_CANDIDATES)
    service_says_medicine = df[SERVICE_COLUMN].map(plain_text).str.contains("medic", regex=False)
    if referral_col:
        referral_values = df[referral_col].map(plain_text)
        explicitly_referred = referral_values.isin({"si", "s", "1", "true", "remitido", "remision", "derivado", "derivacion", "enviado", "entregado", "pasado"})
    else:
        explicitly_referred = pd.Series(False, index=df.index)
    df["_caso_medicina"] = service_says_medicine | explicitly_referred
    df["_remision_medicina"] = explicitly_referred
    date_col = find_column(df.columns, APPOINTMENT_DATE_CANDIDATES)
    time_col = find_column(df.columns, APPOINTMENT_TIME_CANDIDATES)
    address_col = find_column(df.columns, ADDRESS_CANDIDATES)
    df["Dirección"] = df[address_col].astype(str).str.strip() if address_col else ""
    df["_fecha_cita"] = df[date_col].map(parse_excel_date) if date_col else pd.Series(pd.NaT, index=df.index)
    df["Fecha de cita"] = df["_fecha_cita"].map(lambda value: value.strftime("%d/%m/%Y") if pd.notna(value) else "")
    if time_col:
        df["Hora"] = df[time_col].map(format_excel_time)
    else:
        df["Hora"] = ""

    df_view = df.copy()
    # Incluye también los casos que todavía no tienen servicio/cita asignados.
    df_view["_servicio_filtro"] = df_view[SERVICE_COLUMN].astype(str).str.strip().replace("", "Sin servicio agendado")
    territory_code_col = next((name for name in ["Cod Territorio", "Cod_Territorio"] if name in df_view.columns), None)
    territory_name_col = "Nombre Territorio" if "Nombre Territorio" in df_view.columns else None
    if territory_code_col:
        codes = df_view[territory_code_col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        names = df_view[territory_name_col].astype(str).str.strip() if territory_name_col else pd.Series("", index=df_view.index)
        df_view["_territorio_filtro"] = [
            f"{code} · {TERRITORY_DIRECTORY.get(code, name)}" if code else (name or "Sin territorio registrado")
            for code, name in zip(codes, names)
        ]
    elif territory_name_col:
        names = df_view[territory_name_col].astype(str).str.strip()
        df_view["_territorio_filtro"] = names.replace("", "Sin territorio registrado")
    else:
        df_view["_territorio_filtro"] = "Sin territorio registrado"
    with st.sidebar:
        st.caption(source_label)
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()
        page = st.radio("Ir a", ["Personas", "Profesionales y citas", "Usuarios capitados", "Seguimiento"], index=3, key="main_navigation_v2", label_visibility="collapsed")
        st.divider()
        st.header("Filtros de la vista")
        service_options = sorted(df_view["_servicio_filtro"].unique())
        official_territories = [f"{code} · {name}" for code, name in TERRITORY_DIRECTORY.items()]
        observed_territories = sorted(set(df_view["_territorio_filtro"].unique()) - set(official_territories))
        territory_options = official_territories + observed_territories
        if st.button("Mostrar todos los casos", use_container_width=True):
            st.session_state["filter_services"] = service_options
            st.session_state["filter_territories"] = territory_options
            st.session_state["filter_search"] = ""
        selected_services = st.multiselect("Servicio (incluye casos sin asignar)", service_options, default=service_options, key="filter_services")
        selected_territories = st.multiselect("Territorio", territory_options, default=territory_options, key="filter_territories")
        search = st.text_input("Filtrar por nombre o código", key="filter_search")
        if st.button("Actualizar desde Excel", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        st.caption(f"Base completa leída: {len(df):,} casos")
    st.caption(f"Fuente: {df['Hoja_origen'].nunique()} pestañas del libro · {source_label}")
    if len(df) != EXPECTED_CASE_COUNT:
        st.warning(f"El archivo leído tiene {len(df):,} casos; el total esperado en tu base es {EXPECTED_CASE_COUNT:,}. Revisa que hayas cargado la versión actualizada del Excel.")

    mask = pd.Series(True, index=df_view.index)
    if selected_services:
        mask &= df_view["_servicio_filtro"].isin(selected_services)
    else:
        mask &= False
    if selected_territories:
        mask &= df_view["_territorio_filtro"].isin(selected_territories)
    else:
        mask &= False
    if search.strip():
        needle = norm(search)
        codes = df_view.get("Cod_Persona", pd.Series("", index=df_view.index)).map(norm)
        mask &= df_view["Paciente"].map(norm).str.contains(needle, regex=False) | codes.str.contains(needle, regex=False)
    filtered = df_view.loc[mask].copy()
    appointments_all = filtered.loc[filtered["_fecha_cita"].notna()].copy()
    # Una cita solo cuenta cuando tiene fecha, hora, servicio válido y profesional identificable.
    known_professionals = {norm(person) for people in PROFESSIONAL_DIRECTORY.values() for person in people}
    appointments = appointments_all.loc[
        appointments_all["Servicio de agenda"].isin(APPOINTMENT_SERVICES)
        & appointments_all["Hora"].astype(str).str.strip().ne("")
        & appointments_all["Profesional"].map(norm).isin(known_professionals)
    ].copy()
    no_appointment = filtered.loc[filtered["_fecha_cita"].isna()].copy()

    total, scheduled = len(filtered), len(appointments)
    followup_col = find_column(filtered.columns, FOLLOWUP_CANDIDATES)
    if followup_col:
        followup = filtered[followup_col].map(norm)
    else:
        followup = pd.Series("", index=filtered.index)
    pending_values = {"", "nan", "none", "nat", "pendiente", "por llamar", "sin llamar", "pendiente de llamar", "pendiente de gestionar"}
    has_scheduled_visit = filtered.index.isin(appointments.index)
    pending_status_values = {"pendiente", "por remitir", "pendiente de remision", "por remisionar", "si", "yes", "x", "1", "1.0", "true", "verdadero", "remitir", "marcado", "checked", "☑", "✓", "✔"}
    generic_referral_col = find_column(filtered.columns, GENERIC_REFERRAL_STATUS_CANDIDATES)
    service_for_referral_col = find_column(filtered.columns, SERVICE_CANDIDATES)
    referral_counts = {}
    any_referral_marked = pd.Series(False, index=filtered.index)
    for referral_service, candidates in REFERRAL_STATUS_CANDIDATES.items():
        status_col = find_column(filtered.columns, candidates)
        if status_col:
            service_marked = filtered[status_col].map(plain_text).isin(pending_status_values)
            referral_counts[referral_service] = int(service_marked.sum())
            any_referral_marked |= service_marked
        elif generic_referral_col and service_for_referral_col:
            status_pending = filtered[generic_referral_col].map(plain_text).isin(pending_status_values)
            service_text = filtered[service_for_referral_col].map(plain_text)
            service_token = {"Medicina": "medic", "Psicología": "psicolog", "Enfermería": "enfermer", "Nutrición": "nutri"}[referral_service]
            service_marked = status_pending & service_text.str.contains(service_token, regex=False)
            referral_counts[referral_service] = int(service_marked.sum())
            any_referral_marked |= service_marked
        else:
            referral_counts[referral_service] = 0
    managed_mask = (~followup.isin(pending_values)) | has_scheduled_visit | filtered["_remision_medicina"] | any_referral_marked
    managed_count = int(managed_mask.sum())
    pending_count = int((~managed_mask).sum())
    st.caption(f"{source_label} · Actualizado: {datetime.now(BOGOTA_TZ).strftime('%d/%m/%Y %I:%M:%S %p')}")

    if page == "Personas":
        st.subheader("👥 Personas")
        st.caption("Busca un caso en las pestañas del Excel cargado. La búsqueda no modifica el libro.")
        doc_type_col = find_column(filtered.columns, ["Tipo_Documento", "Tipo de documento", "Tipo Documento"])
        doc_number_col = find_column(filtered.columns, ["N°_Documento", "N° Documento", "No. Documento", "Numero_Documento", "Número de documento", "Documento", "Numero Documento"])
        with st.form("person_lookup"):
            a, b, c = st.columns([1, 1.4, 0.5])
            if doc_type_col:
                types = sorted(v for v in filtered[doc_type_col].astype(str).str.strip().unique() if v)
                chosen_type = a.selectbox("Tipo de documento", ["Todos", *types])
            else:
                chosen_type = "Todos"
                a.info("El Excel no tiene una columna de tipo de documento.")
            document = b.text_input("Número de documento")
            submitted = c.form_submit_button("🔎 Buscar", use_container_width=True)
        if submitted:
            if not document.strip():
                st.warning("Escribe el número de documento para buscar.")
            elif not doc_number_col:
                st.error("No encontré una columna de número de documento en el Excel.")
            else:
                found = filtered[filtered[doc_number_col].astype(str).str.strip().str.casefold() == document.strip().casefold()].copy()
                if chosen_type != "Todos" and doc_type_col:
                    found = found[found[doc_type_col].astype(str).str.strip() == chosen_type]
                if found.empty:
                    st.info("No encontré coincidencias con esos datos y los filtros seleccionados.")
                else:
                    st.success(f"Encontré {len(found)} registro(s).")
                    show_cols = [c for c in ["Paciente", doc_type_col, doc_number_col, "Cod_Persona", "Hoja_origen", "Cod Territorio", "Nombre Territorio", SERVICE_COLUMN, "Fecha de cita", "Hora", "Dirección"] if c and c in found.columns]
                    st.dataframe(found[show_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Ingresa el documento y pulsa Buscar para consultar a una persona.")

    elif page == "Profesionales y citas":
        st.subheader("🩺 Profesionales y citas")
        st.caption("Citas domiciliarias que ya estén registradas en el Excel para Psicología, Enfermería o Nutrición.")
        providers = pd.DataFrame([(service, person) for service, people in PROFESSIONAL_DIRECTORY.items() for person in people], columns=["Servicio", "Profesional"])
        st.markdown("**Directorio profesional**")
        st.dataframe(providers, use_container_width=True, hide_index=True)
        st.markdown("**Agenda registrada**")
        if appointments.empty:
            st.info("Aún no hay citas registradas. Cuando agregues en Excel el paciente, dirección, profesional, fecha y hora, aparecerán aquí.")
        else:
            agenda_cols = [c for c in ["Hoja_origen", "Fecha de cita", "Hora", "Paciente", "Cod_Persona", "Dirección", "Servicio de agenda", "Profesional", APPOINTMENT_PLACE_COLUMN, "Agendamiento_Promotor"] if c in appointments.columns]
            appointments = appointments.sort_values("_fecha_cita")
            st.dataframe(appointments[agenda_cols], use_container_width=True, hide_index=True)
            st.download_button("Descargar agenda", appointments[agenda_cols].to_csv(index=False).encode("utf-8-sig"), "agenda_citas.csv", "text/csv")
        st.info("Los casos de Medicina se cuentan aparte para remitirlos a la persona encargada; no aparecen como citas gestionadas por ti.")

    elif page == "Usuarios capitados":
        st.subheader("👥 Usuarios capitados")
        st.caption("Se identifica la EPS en EAPB/EPS y el régimen subsidiado en EAPB o en la columna Régimen.")
        eapb_cols = list(dict.fromkeys(col for candidate in EAPB_CANDIDATES if (col := find_column(filtered.columns, [candidate]))))
        regime_col = find_column(filtered.columns, REGIME_CANDIDATES)
        address_display_col = find_column(filtered.columns, ADDRESS_DISPLAY_CANDIDATES)
        pcp_col = find_column(filtered.columns, PCP_CANDIDATES)
        if not eapb_cols:
            st.warning("No encontré una columna de EAPB/EPS en el Excel para identificar a los usuarios capitados.")
        else:
            def insurer_values(frame: pd.DataFrame) -> pd.Series:
                values = frame[eapb_cols].replace(r"^\s*$", pd.NA, regex=True).bfill(axis=1).iloc[:, 0]
                return values.fillna("").astype(str).str.strip()

            def capitated_rows(frame: pd.DataFrame) -> pd.Series:
                insurer = insurer_values(frame)
                normalized_eapb = insurer.map(plain_text)
                if regime_col:
                    is_subsidized = normalized_eapb.str.contains("subsid", regex=False) | frame[regime_col].map(plain_text).str.contains("subsid", regex=False)
                else:
                    is_subsidized = normalized_eapb.str.contains("subsid", regex=False)
                has_target_eapb = normalized_eapb.str.contains("capital", regex=False) | normalized_eapb.str.contains("sanitas", regex=False)
                return has_target_eapb & is_subsidized

            all_capitated_mask = capitated_rows(df_view)
            filtered_capitated_mask = capitated_rows(filtered)
            capitated = filtered.loc[filtered_capitated_mask].copy()
            capitated["Asegurador capitado"] = insurer_values(capitated)
            st.metric("Usuarios capitados", len(capitated))
            if capitated.empty:
                if all_capitated_mask.any():
                    st.warning(f"Sí hay {int(all_capitated_mask.sum())} usuario(s) capitado(s) en el Excel, pero los filtros actuales los están ocultando. Pulsa ‘Mostrar todos los casos’ en el menú lateral.")
                else:
                    eapb_examples = sorted(v for v in insurer_values(df_view).unique() if v)[:8]
                    regime_examples = sorted(v for v in df_view[regime_col].astype(str).str.strip().unique() if v)[:8] if regime_col else []
                    st.warning("No encontré esa usuaria como Capital/Sanitas subsidiado en el archivo cargado. Verifica que cargaste el Excel correcto desde la barra lateral.")
                    if eapb_examples:
                        st.caption("Valores de EPS que sí aparecen en la base: " + ", ".join(eapb_examples))
                    if regime_examples:
                        st.caption("Valores de régimen que sí aparecen en la base: " + ", ".join(regime_examples))
            else:
                display_cols = [c for c in ["Paciente", "Tipo_Documento", "N°_Documento", "Cod_Persona", "Asegurador capitado", regime_col, address_display_col, pcp_col, "Cod Territorio", "Nombre Territorio", "Hoja_origen"] if c and c in capitated.columns]
                st.dataframe(capitated[display_cols], use_container_width=True, hide_index=True)
                st.download_button("Descargar usuarios capitados", capitated[display_cols].to_csv(index=False).encode("utf-8-sig"), "usuarios_capitados.csv", "text/csv")

    else:
        st.subheader("📊 Seguimiento")
        k1, k2, k3 = st.columns(3)
        k1.metric("Casos en el libro", len(df))
        k2.metric("Casos en la vista", total)
        k3.metric("Pendientes de gestión", pending_count)
        referrals = st.columns(4)
        for metric, referral_service in zip(referrals, ["Medicina", "Psicología", "Enfermería", "Nutrición"]):
            metric.metric(f"Por remitir a {referral_service}", referral_counts[referral_service])
        st.caption("Marca Sí, X o una casilla verdadera en la columna Remitir a cada servicio. Los casos sin marcar o con No no se cuentan.")
        st.subheader("Casos por hoja")
        by_sheet = filtered.groupby("Hoja_origen").size().rename("Casos").reset_index()
        st.dataframe(by_sheet, use_container_width=True, hide_index=True)
        st.subheader("Casos por territorio")
        by_territory = filtered.groupby("_territorio_filtro").size().rename("Casos").reset_index().rename(columns={"_territorio_filtro": "Territorio"})
        st.dataframe(by_territory, use_container_width=True, hide_index=True)
        st.caption("‘Sin fecha de cita’ no significa automáticamente ‘sin llamar’; usa Estado_Gestion para registrar el seguimiento real.")

        st.subheader("Registros sin fecha de cita")
        st.caption("Lista de seguimiento sugerida. La ausencia de fecha no confirma por sí sola que falte una llamada.")
        cols = [c for c in ["Hoja_origen", "Paciente", "Cod_Persona", "Subred", "Zona", "Localidad", "Barrio", SERVICE_COLUMN, "Detalle de alertas", STATUS_COLUMN] if c in no_appointment.columns]
        st.dataframe(no_appointment[cols], use_container_width=True, hide_index=True)
        st.download_button("Descargar lista", no_appointment[cols].to_csv(index=False).encode("utf-8-sig"), "casos_sin_cita.csv", "text/csv")

    st.caption("Solo lectura: el tablero no modifica el libro. Carga el Excel actualizado desde la barra lateral para que se refleje aquí.")


dashboard()
