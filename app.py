import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
import graphviz
import io

# Set page layout
st.set_page_config(page_title="Excel Named Range Visualizer", layout="wide")

# Initialize OpenAI client with API key from Streamlit secrets
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# --- Extract named references from Excel workbook ---
@st.cache_data(show_spinner=False)
def extract_named_references(_wb):
    named_refs = {}
    for name in _wb.defined_names:
        defined_name = _wb.defined_names[name]
        if defined_name.attr_text and not defined_name.is_external:
            dests = list(defined_name.destinations)
            for sheet_name, ref in dests:
                named_refs[defined_name.name] = {
                    "sheet": sheet_name,
                    "ref": ref,
                    "formula": None
                }
                try:
                    sheet = _wb[sheet_name]
                    cell_ref = ref.split('!')[-1]
                    cell = sheet[cell_ref]
                    if cell.data_type == 'f':
                        named_refs[defined_name.name]["formula"] = cell.value
                except Exception:
                    pass
    return named_refs

# --- Detect dependencies between named references ---
@st.cache_data(show_spinner=False)
def find_dependencies(named_refs):
    dependencies = {}
    for name, info in named_refs.items():
        formula = info.get("formula", "")
        if formula:
            formula = formula.upper()
            deps = [other for other in named_refs if other != name and other.upper() in formula]
            dependencies[name] = deps
        else:
            dependencies[name] = []
    return dependencies

# --- Create a Graphviz dependency graph ---
def create_dependency_graph(dependencies):
    dot = graphviz.Digraph()
    for ref in dependencies:
        dot.node(ref)
    for ref, deps in dependencies.items():
        for dep in deps:
            dot.edge(dep, ref)
    return dot

# --- Call OpenAI to explain formulas and convert them to Python ---
@st.cache_data(show_spinner=False)
def call_openai(prompt, max_tokens=100):
    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=max_tokens
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"(Error: {e})"

# --- Generate AI documentation + Python translation for each named ref ---
@st.cache_data(show_spinner=False)
def generate_ai_outputs(named_refs):
    results = []
    for name, info in named_refs.items():
        excel_formula = info.get("formula", "")
        if not excel_formula:
            doc = "No formula."
            py = ""
        else:
            doc_prompt = f"Explain what the following Excel formula does:\n{excel_formula}"
            py_prompt = f"Translate this Excel formula into a clean, readable Python expression:\n{excel_formula}"
            doc = call_openai(doc_prompt, max_tokens=100)
            py = call_openai(py_prompt, max_tokens=100)

        results.append({
            "Named Reference": name,
            "AI Documentation": doc,
            "Excel Formula": excel_formula,
            "Python Formula": py,
        })
    return results

# --- Streamlit App UI ---
st.title("📊 Excel Named Range Dependency Viewer with AI")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(filename=io.BytesIO(uploaded_file.read()), data_only=False)

        st.subheader("📌 Named References Found")
        named_refs = extract_named_references(_wb=wb)
        st.json(named_refs)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_dependencies(named_refs)
        dot = create_dependency_graph(dependencies)
        st.graphviz_chart(dot)

        st.subheader("🧠 AI-Generated Documentation and Python Translation")
        with st.spinner("Asking GPT for documentation and conversions..."):
            table_rows = generate_ai_outputs(named_refs)
            st.dataframe(table_rows, use_container_width=True)

    except Exception as e:
        st.error(f"⚠️ Failed to process file: {e}")
else:
    st.info("Please upload a `.xlsx` file to begin.")
