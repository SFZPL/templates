import os
import io
import xmlrpc.client
import streamlit as st
from dotenv import load_dotenv
from docx import Document
import datetime
from typing import Tuple, Optional, List, Dict, Any

# Load environment variables
load_dotenv()

# === Odoo Credentials from Environment Variables ===
ODOO_URL = os.getenv("ODOO_URL", "https://prezlab-staging-17999869.dev.odoo.com")
ODOO_DB = os.getenv("ODOO_DB", "prezlab-staging-17999869")
ODOO_USERNAME = os.getenv("ODOO_USERNAME", "omar.elhasan@prezlab.com")
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD", "1234")

# --- Odoo Connection using caching for robustness and speed ---
@st.cache_resource(show_spinner=False)
def get_odoo_connection() -> Tuple[Optional[int], Optional[xmlrpc.client.ServerProxy]]:
    """
    Authenticates with Odoo and returns the user id and models proxy.
    """
    try:
        common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
        uid = common.authenticate(ODOO_DB, ODOO_USERNAME, ODOO_PASSWORD, {})
        if not uid:
            st.error("Failed to authenticate with Odoo. Check credentials and database name.")
            return None, None
        models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
        return uid, models
    except Exception as e:
        st.error(f"Error connecting to Odoo: {e}")
        return None, None

def get_employee_fields(models: xmlrpc.client.ServerProxy, uid: int) -> List[str]:
    """
    Retrieves available fields for the hr.employee model.
    """
    try:
        fields = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'hr.employee', 'fields_get', [], {'attributes': ['type']}
        )
        available_fields = list(fields.keys())
        return available_fields
    except Exception as e:
        st.error(f"Error retrieving employee fields: {e}")
        return []

def get_arabic_name(employee: Dict[str, Any]) -> str:
    """
    Attempts to retrieve the Arabic name from multiple possible fields.
    """
    possible_keys = ["x_studio_employee_arabic_name", "Employee Arabic Name", "arabic_name"]
    for key in possible_keys:
        name = employee.get(key, "").strip()
        if name:
            return name
    return employee.get("name", "").strip()

def get_employee_by_phone(models: xmlrpc.client.ServerProxy, uid: int, phone_number: str) -> Optional[Dict[str, Any]]:
    """
    Searches for an employee by phone number and returns processed employee data.
    """
    try:
        phone_number = phone_number.strip()
        available_fields = get_employee_fields(models, uid)

        # Build search domain using available phone fields.
        if "mobile_phone" in available_fields and "work_phone" in available_fields:
            search_domain = ['|', ('mobile_phone', '=', phone_number), ('work_phone', '=', phone_number)]
        elif "mobile_phone" in available_fields:
            search_domain = [('mobile_phone', '=', phone_number)]
        elif "work_phone" in available_fields:
            search_domain = [('work_phone', '=', phone_number)]
        else:
            st.error("Neither 'mobile_phone' nor 'work_phone' exist in Odoo.")
            return None

        # Define fields to read; intentionally using create_date as the joining date.
        fields_to_read = ['id', 'name', 'job_title', 'create_date', 'x_studio_employee_arabic_name']
        if "mobile_phone" in available_fields:
            fields_to_read.append("mobile_phone")
        if "work_phone" in available_fields:
            fields_to_read.append("work_phone")

        employee_ids = models.execute_kw(ODOO_DB, uid, ODOO_PASSWORD, 'hr.employee', 'search', [search_domain])
        if not employee_ids:
            st.warning("No employee found with the provided phone number.")
            return None

        if len(employee_ids) > 1:
            st.warning("Multiple employees found; using the first match.")

        employee_data = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD, 'hr.employee', 'read',
            [employee_ids], {'fields': fields_to_read}
        )
        if not employee_data:
            st.warning("Employee data retrieval failed.")
            return None

        employee = employee_data[0]

        try:
            contracts = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD, 'hr.contract', 'search_read',
                [[('employee_id', '=', employee['id'])]],
                {'fields': ['wage'], 'limit': 1}
            )
        except xmlrpc.client.Fault as fault:
            st.warning("User does not have access to hr.contract records. Default wage will be used.")
            contracts = []
        wage = contracts[0].get('wage', 0.0) if contracts else 0.0

        join_date_raw = employee.get('create_date', '')
        join_date_str = ""
        if join_date_raw:
            try:
                join_date_dt = datetime.datetime.strptime(join_date_raw.split(" ")[0], "%Y-%m-%d")
                join_date_str = join_date_dt.strftime("%d/%m/%Y")
            except Exception:
                join_date_str = join_date_raw

        arabic_name = get_arabic_name(employee)

        return {
            'id': employee.get('id', ''),
            'name': employee.get('name', '').strip(),
            'first_name': employee.get('name', '').split()[0] if employee.get('name') else '',
            'job_title': employee.get('job_title', '').strip(),
            'phone': employee.get('mobile_phone', '').strip() or employee.get('work_phone', '').strip(),
            'wage': wage,
            'joining_date': join_date_str,
            'arabic_name': arabic_name
        }
    except Exception as e:
        st.error(f"Error retrieving employee data: {e}")
        return None

def replace_text_in_runs(runs: List[Any], placeholder: str, replacement: str) -> None:
    """
    Replaces text in Word runs while preserving formatting.
    """
    full_text = "".join(run.text for run in runs)
    new_text = full_text.replace(placeholder, replacement)
    if new_text != full_text:
        for run in runs:
            run.text = ""
        runs[0].text = new_text

def fill_template(template_path: str, employee_data: Dict[str, Any], is_arabic: bool = False) -> Optional[bytes]:
    """
    Loads a Word document template, replaces placeholders with employee data,
    and returns the modified document as bytes.
    """
    if not os.path.exists(template_path):
        st.error(f"Template file not found: {template_path}")
        return None

    try:
        doc = Document(template_path)
    except Exception as e:
        st.error(f"Error loading document: {e}")
        return None

    current_date = datetime.date.today().strftime("%d/%m/%Y")

    placeholders = {
        "(Current Date)": current_date,
        "(First and Last Name)": employee_data['name'],
        "(First Name)": employee_data['first_name'],
        "(Position)": employee_data['job_title'],
        "(Salary)": str(employee_data['wage']),
        "(DD/MM/YYYY)": employee_data['joining_date'],
        "(Country)": employee_data.get('country', ''),
        "(Start Date)": employee_data.get('start_date', ''),
        "(End Date)": employee_data.get('end_date', ''),
        "(الاسم الكامل)": employee_data.get("arabic_name", employee_data['name']) if is_arabic else employee_data['name'],
        "(بلد الوجهة)": employee_data.get('country', ''),
        "(تاريخ البداية)": employee_data.get('start_date', ''),
        "(تاريخ النهاية)": employee_data.get('end_date', '')
    }

    for para in doc.paragraphs:
        for key, value in placeholders.items():
            replace_text_in_runs(para.runs, key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, value in placeholders.items():
                        replace_text_in_runs(para.runs, key, value)
                        
    for section in doc.sections:
        for para in section.header.paragraphs:
            for key, value in placeholders.items():
                replace_text_in_runs(para.runs, key, value)
        for para in section.footer.paragraphs:
            for key, value in placeholders.items():
                replace_text_in_runs(para.runs, key, value)

    output_stream = io.BytesIO()
    doc.save(output_stream)
    return output_stream.getvalue()

# Configure page settings.
st.set_page_config(
    page_title="Employment Letter Generator",
    page_icon=":briefcase:",
    layout="centered"
)

# Minimal CSS for a clean appearance.
st.markdown(
    """
    <style>
    .main { padding: 2rem; }
    .stButton>button { background-color: #2e7bcf; color: white; border-radius: 5px; }
    </style>
    """,
    unsafe_allow_html=True
)

# Template options with new paths.
template_options = {
    "Employment letter - Arabic": r"C:\Users\Geeks\Desktop\Programming_Files\Letters\Employment Letter - ARABIC.docx",
    "Employment letter": r"C:\Users\Geeks\Desktop\Programming_Files\Letters\Employment Letter .docx",
    "Employment letter to embassies": r"C:\Users\Geeks\Desktop\Programming_Files\Letters\Employment Letter to Embassies.docx",
    "Experience letter": r"C:\Users\Geeks\Desktop\Programming_Files\Letters\Experience Letter.docx"
}

def main() -> None:
    st.title("Employment Letter Generator 🚀")
    st.markdown("Please fill in the details below to generate the employment letter. ✨")

    with st.form("letter_form", clear_on_submit=True):
        phone_number = st.text_input("Employee Mobile Number 📱")
        selected_template = st.selectbox("Select Template", list(template_options.keys()))
        template_path = template_options[selected_template]

        # Show travel details only for embassy letters.
        if selected_template == "Employment letter to embassies":
            country = st.text_input("Country Name 🌍")
            start_date = st.date_input("Travel Start Date 📆")
            end_date = st.date_input("Travel End Date 📆")
        else:
            country, start_date, end_date = "", None, None

        submitted = st.form_submit_button("Generate Letter ✨")
        if submitted:
            uid, models = get_odoo_connection()
            if not uid:
                return

            employee_data = get_employee_by_phone(models, uid, phone_number)
            if not employee_data:
                st.error("Could not retrieve employee data.")
                return

            if selected_template == "Employment letter to embassies":
                employee_data['country'] = country.strip()
                employee_data['start_date'] = start_date.strftime("%d/%m/%Y") if start_date else ""
                employee_data['end_date'] = end_date.strftime("%d/%m/%Y") if end_date else ""
            else:
                employee_data['country'] = ""
                employee_data['start_date'] = ""
                employee_data['end_date'] = ""

            st.markdown("### Employee Details")
            st.write(f"**Name:** {employee_data.get('name', '')}")
            st.write(f"**Job Title:** {employee_data.get('job_title', '')}")
            st.write(f"**Joining Date:** {employee_data.get('joining_date', '')}")
            st.write(f"**Wage:** {employee_data.get('wage', '')}")

            safe_name = employee_data['name'].replace(' ', '_')
            filename = f"Employment_Letter_{safe_name}.docx"
            doc_bytes = fill_template(template_path, employee_data, is_arabic=(selected_template == "Employment letter - Arabic"))

            if not doc_bytes:
                st.error("Failed to generate the document.")
            else:
                st.session_state.letter_bytes = doc_bytes
                st.session_state.letter_filename = filename
                st.success("Employment Letter Generated Successfully! 🎉")

    if st.session_state.get("letter_bytes"):
        st.download_button(
            label="Download Employment Letter 📄",
            data=st.session_state.letter_bytes,
            file_name=st.session_state.letter_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

if __name__ == "__main__":
    main()
