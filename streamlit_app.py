import streamlit as st
import pandas as pd
import io
# We only need the graphviz library to CREATE the graph object
from graphviz import Digraph

# --- Helper Functions (Unchanged) ---


@st.cache_data
def load_data(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        required_columns = ['Code article PF', 'Designation PF',
                            'Code Article', 'Désignation article', 'Quantité pesée']
        if not all(col in df.columns for col in required_columns):
            st.error(
                f"Error: Missing required columns. Ensure the file has: {required_columns}")
            return None
        return df
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None


def get_bom(df, selected_pf_code):
    bom_df = df[df['Code article PF'] == selected_pf_code].copy()
    if bom_df.empty:
        return pd.DataFrame(), 0
    bom_agg = bom_df.groupby(['Code Article', 'Désignation article'])[
        'Quantité pesée'].sum().reset_index()
    total_weight = bom_agg['Quantité pesée'].sum()
    if total_weight > 0:
        bom_agg['Percentage (%)'] = (
            bom_agg['Quantité pesée'] / total_weight) * 100
    else:
        bom_agg['Percentage (%)'] = 0
    return bom_agg, total_weight


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='BOM')
    processed_data = output.getvalue()
    return processed_data

# --- Recursive BOM Explosion Function (Unchanged) ---


def explode_bom_recursive(df, product_code, required_quantity=1.0, level=0, master_pf_list=None):
    if master_pf_list is None:
        master_pf_list = df['Code article PF'].unique()
    direct_bom, batch_size = get_bom(df, product_code)
    if direct_bom.empty or batch_size == 0:
        return pd.DataFrame()
    final_bom_list = []
    for _, row in direct_bom.iterrows():
        component_code, component_designation, component_percentage = row[
            'Code Article'], row['Désignation article'], row['Percentage (%)']
        calculated_quantity = required_quantity * \
            (component_percentage / 100.0)
        if component_code in master_pf_list:
            sub_bom = explode_bom_recursive(
                df, component_code, calculated_quantity, level + 1, master_pf_list)
            final_bom_list.append(sub_bom)
        else:
            final_bom_list.append(pd.DataFrame([{'Level': level + 1, 'Code Article': component_code,
                                  'Désignation article': component_designation, 'Required Quantity': calculated_quantity, 'Type': 'Raw Material'}]))
    if not final_bom_list:
        return pd.DataFrame()
    full_bom = pd.concat(final_bom_list, ignore_index=True)
    if level == 0:
        return full_bom.groupby(['Code Article', 'Désignation article', 'Type'])['Required Quantity'].sum().reset_index()
    return full_bom

# --- Graph Generation Functions (Unchanged) ---


def add_nodes_edges_recursive(df, product_code, dot, master_pf_list, processed_nodes=None):
    if processed_nodes is None:
        processed_nodes = set()
    if product_code in processed_nodes:
        return
    processed_nodes.add(product_code)
    direct_bom, _ = get_bom(df, product_code)
    if direct_bom.empty:
        return
    product_name = df[df['Code article PF'] ==
                      product_code]['Designation PF'].iloc[0]
    parent_id = str(product_code)
    dot.node(parent_id, f"{product_name}\n({product_code})",
             shape='box', style='filled', fillcolor='lightblue')
    for _, row in direct_bom.iterrows():
        child_code, child_name, percentage = row['Code Article'], row[
            'Désignation article'], row['Percentage (%)']
        child_id = str(child_code)
        if child_code in master_pf_list:
            dot.node(child_id, f"{child_name}\n({child_code})",
                     shape='box', style='filled', fillcolor='lightgray')
            add_nodes_edges_recursive(
                df, child_code, dot, master_pf_list, processed_nodes=processed_nodes)
        else:
            dot.node(child_id, f"{child_name}\n({child_code})",
                     shape='ellipse', style='filled', fillcolor='honeydew')
        dot.edge(parent_id, child_id, label=f"{percentage:.2f}%")


def generate_bom_graph(df, selected_pf_code):
    """Creates the complete BOM graph for a given product."""
    dot = Digraph(f'BOM for {selected_pf_code}')

    # This is the line that has been changed to reorient the graph
    dot.attr(rankdir='LR', splines='ortho')

    dot.attr('node', shape='box', style='rounded')
    master_pf_list = df['Code article PF'].unique()
    add_nodes_edges_recursive(df, selected_pf_code, dot, master_pf_list)
    return dot


def explode_optimized_recursive(df, product_code, required_quantity, min_weighable_qty, master_pf_list=None):
    """
    Recursively explodes a BOM, but stops exploding a sub-assembly if any of its
    children would fall below the minimum weighable quantity.
    """
    if master_pf_list is None:
        master_pf_list = df['Code article PF'].unique()

    # BASE CASE: The current item is a raw material, so we just return it.
    if product_code not in master_pf_list:
        return pd.DataFrame([{
            'Code Article': product_code,
            'Désignation article': df[df['Code Article'] == product_code]['Désignation article'].iloc[0],
            'Required Quantity': required_quantity,
            'Type': 'Raw Material'
        }])

    # --- DECISION LOGIC: Check if we *can* explode this SFG ---
    direct_bom, batch_size = get_bom(df, product_code)

    if direct_bom.empty or batch_size == 0:
        # If this SFG has no formula, treat it as a single item
        product_name = df[df['Code article PF'] ==
                          product_code]['Designation PF'].iloc[0]
        return pd.DataFrame([{'Code Article': product_code, 'Désignation article': product_name, 'Required Quantity': required_quantity, 'Type': 'Semi-Finished Good (No Formula)'}])

    # Calculate required quantities of all direct children
    direct_bom['child_required_qty'] = required_quantity * \
        (direct_bom['Percentage (%)'] / 100.0)

    # Check if ANY child would be below the minimum threshold
    can_explode = direct_bom['child_required_qty'].min() >= min_weighable_qty

    # --- EXECUTION ---
    if can_explode:
        # If we can explode, recurse for each child
        final_bom_list = []
        for _, row in direct_bom.iterrows():
            sub_bom = explode_optimized_recursive(
                df,
                product_code=row['Code Article'],
                required_quantity=row['child_required_qty'],
                min_weighable_qty=min_weighable_qty,
                master_pf_list=master_pf_list
            )
            final_bom_list.append(sub_bom)
        return pd.concat(final_bom_list, ignore_index=True)
    else:
        # STOPPING CASE: Cannot explode. Treat this SFG as a single ingredient.
        product_name = df[df['Code article PF'] ==
                          product_code]['Designation PF'].iloc[0]
        return pd.DataFrame([{
            'Code Article': product_code,
            'Désignation article': product_name,
            'Required Quantity': required_quantity,
            'Type': 'Semi-Finished Good (Not Exploded)'
        }])


# --- Main Application UI ---
st.set_page_config(layout="wide", page_title="BOM Analyzer")
st.title("Bill of Materials (BOM) Analyzer")
st.header("1. Upload Production Data File")
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    df = load_data(uploaded_file)
    if df is not None:
        st.success("File uploaded successfully!")
        st.header("2. Select a Finished Good (PF) to Analyze")
        pf_list = df.drop_duplicates(subset=['Code article PF'])
        pf_list['display'] = pf_list['Code article PF'] + \
            " | " + pf_list['Designation PF']
        selected_pf_display = st.selectbox(
            'Select a product:', options=pf_list['display'])
        selected_pf_code = selected_pf_display.split(" | ")[0]

        tab1, tab2, tab3, tab4 = st.tabs(
            ["Single-Level BOM", "Full Explosion", "BOM Flowchart", "Optimized Explosion"])

        with tab1:
            st.subheader(f"Direct Components for: {selected_pf_display}")
            bom_df, total_weight = get_bom(df, selected_pf_code)
            if not bom_df.empty:
                st.info(
                    f"The BOM is based on a produced batch size of **{total_weight:.2f}** units.")
                st.dataframe(bom_df.style.format(
                    {'Quantité pesée': '{:.4f}', 'Percentage (%)': '{:.2f}%'}))
            else:
                st.warning("No BOM data found for this product.")
        with tab2:
            st.subheader("Fully Exploded BOM (All Raw Materials)")
            target_quantity = st.number_input(
                "Enter target production quantity:", min_value=1.0, value=100.0, step=10.0)
            full_bom_df = explode_bom_recursive(
                df, selected_pf_code, target_quantity)
            if not full_bom_df.empty:
                st.write(
                    f"Total raw materials required for **{target_quantity}** units of the final product:")
                st.dataframe(full_bom_df.style.format(
                    {'Required Quantity': '{:.4f}'}))
                st.subheader("Scaling Analysis")
                min_weighable_qty = st.number_input(
                    "Enter minimum weighable quantity:", min_value=0.0, value=0.01, step=0.001, format="%.4f", key="min_qty_exploded")
                full_bom_df['Is Feasible?'] = full_bom_df['Required Quantity'] >= min_weighable_qty
                st.dataframe(full_bom_df[['Désignation article', 'Required Quantity', 'Is Feasible?']].style.format({'Required Quantity': '{:.4f}'}).apply(
                    lambda row: ['background-color: #FFCDD2' if not row['Is Feasible?'] else '' for _ in row], axis=1))
            else:
                st.warning(
                    "Could not explode BOM. The selected item might be a raw material itself.")
        with tab3:
            st.subheader("BOM Flowchart Visualization")
            st.info(
                "This chart shows the structure of the product, with semi-finished goods in gray and raw materials in green.")
            try:
                graph = generate_bom_graph(df, selected_pf_code)
                # Use the built-in Streamlit function
                st.graphviz_chart(graph)
            except Exception as e:
                st.error(
                    f"Could not generate graph. Ensure the Graphviz system package is installed. Error: {e}")
        with tab4:  # Or tab3 if you place it before the flowchart
            st.subheader("Optimized Explosion Based on Weighing Constraints")
            st.info(
                "This tool explodes the formula as much as possible. It will stop exploding a semi-finished good "
                "if any of its components would require weighing a quantity smaller than your specified minimum."
            )

            # User Inputs
            col1, col2 = st.columns(2)
            with col1:
                target_qty_optimized = st.number_input(
                    "Enter target production quantity:",
                    min_value=0.01, value=100.0, step=10.0, key="target_optimized"
                )
            with col2:
                min_weighable_qty_optimized = st.number_input(
                    "Enter minimum weighable quantity:",
                    min_value=0.0, value=0.01, step=0.001, format="%.4f", key="min_weighable_optimized"
                )

            if st.button("Calculate Optimized BOM"):
                # Run the new recursive function
                optimized_bom_df = explode_optimized_recursive(
                    df,
                    product_code=selected_pf_code,
                    required_quantity=target_qty_optimized,
                    min_weighable_qty=min_weighable_qty_optimized
                )

                if not optimized_bom_df.empty:
                    # Aggregate results in case the same component appears in different branches
                    final_optimized_bom = optimized_bom_df.groupby(
                        ['Code Article', 'Désignation article', 'Type']
                    )['Required Quantity'].sum().reset_index()

                    st.write(
                        f"Optimized ingredient list for **{target_qty_optimized}** units:")

                    # Apply styling to highlight the non-exploded SFGs
                    def highlight_sfg(row):
                        is_sfg = 'Semi-Finished Good' in row['Type']
                        return ['background-color: #FFF3CD'] * len(row) if is_sfg else [''] * len(row)

                    st.dataframe(final_optimized_bom.style.apply(
                        highlight_sfg, axis=1).format({'Required Quantity': '{:.4f}'}))
                    st.success(
                        "Highlighted rows are semi-finished goods that could not be broken down further and should be weighed as intermediates.")

                else:
                    st.warning("Could not generate an optimized BOM.")
