import streamlit as st
from openpyxl import load_workbook
import graphviz
import io

st.set_page_config(page_title="Excel Named Range Visualizer", layout="wide")

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

# --- Streamlit App ---
st.title("📊 Excel Named Range Dependency Viewer")

uploaded_file = st.file_uploader("Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        wb = load_workbook(filename=io.BytesIO(uploaded_file.read()), data_only=False)

        st.subheader("📌 Named References Found")
        named_refs = extract_named_references(wb)
        st.json(named_refs)

        st.subheader("🔗 Dependency Graph")
        dependencies = find_dependencies(named_refs)
        dot = create_dependency_graph(dependencies)
        st.graphviz_chart(dot)

    except Exception as e:
        st.error(f"⚠️ Failed to process file: {e}")
else:
    st.info("Please upload a `.xlsx` file to begin.")
