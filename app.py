import streamlit as st
from openpyxl import load_workbook
import graphviz
import io
import re

st.set_page_config(page_title="Excel Named Range Visualizer", layout="wide")

# --- Helper Functions ---

def extract_named_references(wb):
    named_refs = {}
    for name in wb.defined_names:
        defined_name = wb.defined_names[name]

        if defined_name.attr_text and not defined_name.is_external:
            dests = list(defined_name.destinations)
            for sheet_name, ref in dests:
                named_refs[defined_name.name] = {
                    "sheet": sheet_name,
                    "ref": ref,
                    "formula": None
                }
                try:
                    sheet = wb[sheet_name]
                    cell_ref = ref.split('!')[-1]
                    cell = sheet[cell_ref]
                    if cell.data_type == 'f':
                        named_refs[defined_name.name]["formula"] = cell.value
                except Exception:
                    pass
    return named_refs

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

def create_dependency_graph(dependencies):
    dot = graphviz.Digraph()
    for ref in dependencies:
        dot.node(ref)
    for ref, deps in dependencies.items():
        for dep in deps:
            dot.edge(dep, ref)  # dep → ref
    return dot

def convert_excel_to_python(formula):
    if not formula:
        return ""
    f = formula.lstrip("=").strip()

    # Basic replacements
    f = re.sub(r"SUM\(([^)]+)\)", r"sum([\1])", f, flags=re.IGNORECASE)
    f = f.replace("^", "**").replace("&", "+")
    f = re.sub(r"\bIF\(([^,]+),([^,]+),([^)]+)\)", r"(\2 if \1 else \3)", f, flags=re.IGNORECASE)
    return f

def generate_doc(formula):
    if not formula:
        return "Constant or no formula."
    if "SUM" in formula.upper():
        return "This calculates the sum of a range of values."
    if "IF" in formula.upper():
        return "This performs a conditional check and returns different values."
    return "Performs a basic arithmetic or reference-based calculation."

# --- Streamlit App ---
st.title("📊 Excel Named Range Dependency Viewer")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(filename=io.BytesIO(uploaded_file.read()), data_only=False)

        # Step 1: Extract named references and formulas
        named_refs = extract_named_references(wb)

        # Step 2: Show Named References JSON
        st.subheader("📌 Named References")
        st.json(named_refs)

        # Step 3: Show Dependency Graph
        dependencies = find_dependencies(named_refs)
        st.subheader("🔗 Dependency Graph")
        dot = create_dependency_graph(dependencies)
        st.graphviz_chart(dot)

        # Step 4: Generate table with docs + formulas
        st.subheader("🧠 Named Reference Details")

        table_rows = []
        for name, info in named_refs.items():
            excel_formula = info.get("formula", "")
            doc = generate_doc(excel_formula)
            python_formula = convert_excel_to_python(excel_formula)
            table_rows.append({
                "Named Reference": name,
                "AI Documentation": doc,
                "Excel Formula": excel_formula,
                "Python Formula": python_formula,
            })

        st.dataframe(table_rows, use_container_width=True)

    except Exception as e:
        st.error(f"⚠️ Failed to process file: {e}")
else:
    st.info("Please upload a `.xlsx` file to begin.")
