import streamlit as st
import pandas as pd
import io
from graphviz import Digraph
import numpy as np

# --- Helper Functions ---


@st.cache_data
def load_data(uploaded_file):
    """Loads and processes the uploaded Excel file."""
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        required_columns = ['Code article PF', 'Designation PF',
                            'Code Article', 'Désignation article', 'Quantité pesée']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Error: Missing required columns: {required_columns}")
            return None
        return df
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None


@st.cache_data
def get_bom(_df, selected_pf_code):
    """Calculates the aggregated SINGLE-LEVEL BOM for a selected finished good."""
    bom_df = _df[_df['Code article PF'] == selected_pf_code].copy()
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


def to_excel_multi_sheet(sheets_dict):
    """Converts a dictionary of DataFrames to an Excel file with multiple sheets."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# --- Core Logic Functions ---


@st.cache_data
def explode_bom_recursive(_df, product_code, required_quantity=1.0, level=0, master_pf_list=None, critical_item_code=None):
    """Recursively explodes a BOM, correctly aggregating total quantities from all sources."""
    if master_pf_list is None:
        master_pf_list = _df['Code article PF'].unique()
    direct_bom, _ = get_bom(_df, product_code)
    if direct_bom.empty:
        return pd.DataFrame()
    final_bom_list = []
    for _, row in direct_bom.iterrows():
        component_code, component_designation, component_percentage = row[
            'Code Article'], row['Désignation article'], row['Percentage (%)']
        calculated_quantity = required_quantity * \
            (component_percentage / 100.0)
        if component_code in master_pf_list:
            sub_bom = explode_bom_recursive(
                _df, component_code, calculated_quantity, level + 1, master_pf_list, critical_item_code)
            final_bom_list.append(sub_bom)
        else:
            source = 'Original Formula' if level == 0 else 'Exploded Sub-Assembly'
            is_critical = (component_code == critical_item_code)
            final_bom_list.append(pd.DataFrame([{'Code Article': component_code, 'Désignation article': component_designation,
                                  'Required Quantity': calculated_quantity, 'Type': 'Raw Material', 'Source': source, 'Is Critical': is_critical}]))
    if not final_bom_list:
        return pd.DataFrame()
    full_bom = pd.concat(final_bom_list, ignore_index=True)
    if level == 0:
        agg_rules = {'Required Quantity': 'sum', 'Is Critical': 'max',
                     'Source': lambda s: ', '.join(sorted(s.unique()))}
        final_bom = full_bom.groupby(
            ['Code Article', 'Désignation article', 'Type']).agg(agg_rules).reset_index()
        return final_bom
    return full_bom


@st.cache_data
def find_formula_critical_component(_df, product_code):
    """Calculates the fully exploded BOM to find the raw material with the lowest effective percentage."""
    exploded_bom_for_1_unit = explode_bom_recursive(_df, product_code, 1.0)
    if exploded_bom_for_1_unit.empty:
        return None
    min_row = exploded_bom_for_1_unit.loc[exploded_bom_for_1_unit['Required Quantity'].idxmin(
    )]
    return {"code": min_row['Code Article'], "designation": min_row['Désignation article'], "percentage": min_row['Required Quantity'] * 100, "source": min_row['Source']}


def calculate_optimized_bom_iterative(df, product_code, target_quantity, min_weighable_qty):
    """Calculates the optimized BOM using an iterative approach that correctly validates every line item at each step."""
    master_pf_list = df['Code article PF'].unique()
    current_bom, _ = get_bom(df, product_code)
    current_bom['Required Quantity'] = target_quantity * \
        (current_bom['Percentage (%)'] / 100.0)
    unexplodable_sfgs = set()
    while True:
        explodable_sfgs = current_bom[current_bom['Code Article'].isin(
            master_pf_list) & ~current_bom['Code Article'].isin(unexplodable_sfgs)].copy()
        if explodable_sfgs.empty:
            break
        sfg_to_explode = explodable_sfgs.loc[explodable_sfgs['Required Quantity'].idxmax(
        )]
        children_bom, _ = get_bom(df, sfg_to_explode['Code Article'])
        children_bom['Required Quantity'] = sfg_to_explode['Required Quantity'] * \
            (children_bom['Percentage (%)'] / 100.0)
        bom_without_sfg = current_bom[current_bom['Code Article']
                                      != sfg_to_explode['Code Article']]
        tentative_bom = pd.concat(
            [bom_without_sfg, children_bom], ignore_index=True)
        is_valid_explosion = tentative_bom['Required Quantity'].min(
        ) >= min_weighable_qty
        if is_valid_explosion:
            current_bom = tentative_bom.groupby(['Code Article', 'Désignation article']).agg({
                'Required Quantity': 'sum'}).reset_index()
        else:
            unexplodable_sfgs.add(sfg_to_explode['Code Article'])

    def get_type(
        code): return 'Semi-Finished Good (Not Exploded)' if code in master_pf_list else 'Raw Material'
    current_bom['Type'] = current_bom['Code Article'].apply(get_type)
    return current_bom

# --- Report Generation Functions ---


def generate_sfg_breakdown_report(df, optimized_bom_df, min_weighable_qty):
    sfg_to_breakdown = optimized_bom_df[optimized_bom_df['Type'].str.contains(
        'Semi-Finished Good')].copy()
    if sfg_to_breakdown.empty:
        return None
    all_breakdowns = []
    for _, sfg_row in sfg_to_breakdown.iterrows():
        sfg_code, total_sfg_quantity_needed = sfg_row['Code Article'], sfg_row['Required Quantity']
        internal_bom_df, _ = get_bom(df, sfg_code)
        if not internal_bom_df.empty:
            header_df = pd.DataFrame([{'Component': f"--- Breakdown for: {sfg_row['Désignation article']} ({sfg_code}) ---", 'Internal Percentage (%)':
                                     f"Total Needed: {total_sfg_quantity_needed:.4f}", 'Tentative Required Quantity': f"Checked Against Min: {min_weighable_qty}", 'Justification': ''}])
            all_breakdowns.append(header_df)
            internal_bom_df['Tentative Required Quantity'] = total_sfg_quantity_needed * (
                internal_bom_df['Percentage (%)'] / 100.0)
            internal_bom_df['Justification'] = np.where(
                internal_bom_df['Tentative Required Quantity'] < min_weighable_qty, 'FAILED: Below Minimum Weight', 'OK')
            report_df = internal_bom_df[['Désignation article', 'Percentage (%)', 'Tentative Required Quantity', 'Justification']].rename(
                columns={'Désignation article': 'Component', 'Percentage (%)': 'Internal Percentage (%)'})
            all_breakdowns.append(report_df)
            all_breakdowns.append(pd.DataFrame([{'Component': ''}]))
    if not all_breakdowns:
        return None
    return pd.concat(all_breakdowns, ignore_index=True)


def generate_comparison_report(initial_bom_df, optimized_bom_df, target_quantity):
    initial_bom_df['Initial Quantity'] = target_quantity * \
        (initial_bom_df['Percentage (%)'] / 100.0)
    initial_bom = initial_bom_df[[
        'Code Article', 'Désignation article', 'Initial Quantity']].copy()
    final_bom = optimized_bom_df[['Code Article', 'Désignation article', 'Required Quantity', 'Type']].rename(
        columns={'Required Quantity': 'Final Quantity'})
    comparison_df = pd.merge(initial_bom, final_bom, on=[
                             'Code Article', 'Désignation article'], how='outer')
    status_conditions = [(comparison_df['Initial Quantity'].notna()) & (comparison_df['Final Quantity'].isna()), (comparison_df['Initial Quantity'].isna()) & (comparison_df['Final Quantity'].notna()), (comparison_df['Initial Quantity'].notna()) & (
        comparison_df['Final Quantity'].notna()) & (comparison_df['Type'].str.contains('Semi-Finished Good')), (comparison_df['Initial Quantity'].notna()) & (comparison_df['Final Quantity'].notna()) & (comparison_df['Type'] == 'Raw Material')]
    status_choices = ['Exploded', 'New from Explosion',
                      'Kept as Intermediate', 'Kept as Raw Material']
    comparison_df['Status'] = np.select(
        status_conditions, status_choices, default='Unchanged')
    report_df = comparison_df[['Code Article', 'Désignation article',
                               'Initial Quantity', 'Final Quantity', 'Status']].fillna(0)
    return report_df.sort_values(by='Status').reset_index(drop=True)


def trace_explosion_paths_recursive(df, product_code, required_quantity, master_pf_list, final_totals_lookup, current_path_names=[]):
    direct_bom, _ = get_bom(df, product_code)
    if direct_bom.empty:
        return pd.DataFrame()
    path_results = []
    for _, row in direct_bom.iterrows():
        child_code, child_name, child_quantity_from_path = row['Code Article'], row[
            'Désignation article'], required_quantity * (row['Percentage (%)'] / 100.0)
        new_path_names = current_path_names + \
            [child_name] if child_code in master_pf_list else current_path_names
        if child_code in master_pf_list:
            deeper_results = trace_explosion_paths_recursive(
                df, child_code, child_quantity_from_path, master_pf_list, final_totals_lookup, new_path_names)
            path_results.append(deeper_results)
        else:
            final_total_quantity = final_totals_lookup.get(child_code, 0)
            path_str = " -> ".join(current_path_names)
            path_results.append(pd.DataFrame([{'Explosion Path': path_str, 'Raw Material Code': child_code, 'Raw Material': child_name,
                                'Contribution from this Path': child_quantity_from_path, 'Final Aggregated Quantity': final_total_quantity}]))
    if not path_results:
        return pd.DataFrame()
    return pd.concat(path_results, ignore_index=True)


def generate_explosion_trace_report(df, initial_bom_df, optimized_bom_df, target_quantity):
    """Generates a context-aware trace of all successfully exploded SFGs without status warnings."""
    master_pf_list = df['Code article PF'].unique()
    initial_bom_df['Initial Quantity'] = target_quantity * \
        (initial_bom_df['Percentage (%)'] / 100.0)
    exploded_sfgs = initial_bom_df[initial_bom_df['Code Article'].isin(
        master_pf_list) & ~initial_bom_df['Code Article'].isin(optimized_bom_df['Code Article'])]
    if exploded_sfgs.empty:
        return None
    final_totals_lookup = optimized_bom_df.set_index('Code Article')[
        'Required Quantity']
    all_traces = []
    for _, sfg_row in exploded_sfgs.iterrows():
        sfg_code = sfg_row['Code Article']
        sfg_quantity = sfg_row['Initial Quantity']
        initial_path_name = df[df['Code article PF']
                               == sfg_code]['Designation PF'].iloc[0]
        trace_df = trace_explosion_paths_recursive(
            df, sfg_code, sfg_quantity, master_pf_list, final_totals_lookup, current_path_names=[initial_path_name])
        all_traces.append(trace_df)
    if not all_traces:
        return None
    full_trace_report = pd.concat(all_traces, ignore_index=True)
    return full_trace_report

# --- Graph Generation Functions ---


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
    dot = Digraph(f'BOM for {selected_pf_code}')
    dot.attr(rankdir='LR', splines='ortho')
    dot.attr('node', shape='box', style='rounded')
    master_pf_list = df['Code article PF'].unique()
    add_nodes_edges_recursive(df, selected_pf_code, dot, master_pf_list)
    return dot


# --- Main Application UI ---
st.set_page_config(layout="wide", page_title="BOM Analyzer")
st.title("Bill of Materials (BOM) Analyzer")
st.header("1. Upload Production Data File")
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    df = load_data(uploaded_file)
    if df is not None:
        st.success("File uploaded successfully!")
        st.header("2. Analyze a Specific Finished Good")
        pf_list_df = df[['Code article PF', 'Designation PF']
                        ].drop_duplicates().copy()
        pf_list_df['display'] = pf_list_df['Code article PF'] + \
            " | " + pf_list_df['Designation PF']
        selected_pf_display = st.selectbox(
            'Select a product:', options=pf_list_df['display'])
        selected_pf_code = selected_pf_display.split(" | ")[0]

        tab1, tab2, tab3, tab4 = st.tabs(
            ["Single-Level BOM", "Full Explosion & Scaling", "Optimized Explosion", "BOM Flowchart"])

        with tab1:
            st.subheader(f"Direct Components for: {selected_pf_display}")
            bom_df, total_weight = get_bom(df, selected_pf_code)
            if not bom_df.empty:
                st.info(
                    f"The displayed BOM is based on an actual produced batch size of **{total_weight:.2f}** units.")
                st.dataframe(bom_df.style.format(
                    {'Quantité pesée': '{:.4f}', 'Percentage (%)': '{:.2f}%'}))
                st.download_button(label="📥 Download as Excel", data=to_excel_multi_sheet(
                    {"BOM": bom_df}), file_name=f"BOM_{selected_pf_code}.xlsx")
            else:
                st.warning("No BOM data found for this product.")

        with tab2:
            st.subheader(f"Scaling Analysis for: {selected_pf_display}")
            critical_item = find_formula_critical_component(
                df, selected_pf_code)
            if critical_item:
                min_weighable_qty = st.number_input(
                    "Enter minimum weighable quantity (g):", value=0.1, step=0.01, format="%.4f", key="min_formula")
                min_percentage = critical_item['percentage'] / 100.0
                min_batch_size = min_weighable_qty / min_percentage if min_percentage > 0 else 0
                st.info(
                    "This analysis identifies the raw material with the lowest total percentage in the fully exploded formula.")
                col1, col2, col3 = st.columns(3)
                col1.metric(label="Critical Item for this Formula",
                            value=critical_item['designation'], help=f"Source: {critical_item['source']}.")
                col2.metric(label="Effective Percentage",
                            value=f"{critical_item['percentage']:.4f}%")
                col3.metric(label="Minimum Possible Batch Size",
                            value=f"{min_batch_size:.4f} g", help=f"Ensures the critical item weighs at least {min_weighable_qty}g.")
                st.subheader("Exploded BOM for a Target Quantity")
                target_quantity = st.number_input(
                    "Enter target production quantity (g):", value=min_batch_size, step=10.0)
                full_bom_df = explode_bom_recursive(
                    df, selected_pf_code, target_quantity, critical_item_code=critical_item['code'])
                if not full_bom_df.empty:
                    st.write(
                        f"Total raw materials for **{target_quantity}g** of the final product:")

                    def highlight_critical(row): return [
                        'background-color: #FFC7CE'] * len(row) if row['Is Critical'] else [''] * len(row)
                    sorted_bom_df = full_bom_df.sort_values(
                        by='Required Quantity').reset_index(drop=True)
                    styled_df = sorted_bom_df.style.apply(
                        highlight_critical, axis=1)
                    st.dataframe(styled_df.format({'Required Quantity': '{:.4f}'}).hide(
                        ['Is Critical', 'Code Article', 'Type'], axis=1), use_container_width=True)
                    st.download_button(label="📥 Download Exploded BOM as Excel", data=to_excel_multi_sheet(
                        {"Exploded BOM": sorted_bom_df.drop(columns=['Is Critical'])}), file_name=f"Exploded_BOM_{selected_pf_code}.xlsx")
            else:
                st.warning("Could not analyze this formula.")

        with tab3:
            st.subheader("Optimized Explosion Based on Weighing Constraints")
            st.info("This intelligent tool explodes the formula as much as possible without violating the minimum weight for any single ingredient.")
            col1, col2 = st.columns(2)
            with col1:
                target_qty_optimized = st.number_input(
                    "Enter target production quantity:", value=100.0, key="target_optimized")
            with col2:
                min_weighable_qty_optimized = st.number_input(
                    "Enter minimum weighable quantity:", value=0.1, format="%.4f", key="min_weighable_optimized")
            if st.button("Calculate Optimized BOM"):
                optimized_bom_df = calculate_optimized_bom_iterative(
                    df, selected_pf_code, target_qty_optimized, min_weighable_qty_optimized)
                if not optimized_bom_df.empty:
                    st.write(
                        f"Optimized ingredient list for **{target_qty_optimized}** units:")

                    def highlight_sfg(row): return ['background-color: #FFF3CD'] * len(
                        row) if 'Semi-Finished Good' in row['Type'] else [''] * len(row)
                    st.dataframe(optimized_bom_df.style.apply(highlight_sfg, axis=1).format(
                        {'Required Quantity': '{:.4f}'}), use_container_width=True)
                    st.success(
                        "Highlighted rows are intermediates that were kept to avoid violating weight constraints.")

                    sheets_to_download = {"Optimized BOM": optimized_bom_df}
                    breakdown_report_df = generate_sfg_breakdown_report(
                        df, optimized_bom_df, min_weighable_qty_optimized)
                    if breakdown_report_df is not None:
                        sheets_to_download["SFG Breakdowns (Justification)"] = breakdown_report_df
                    initial_bom, _ = get_bom(df, selected_pf_code)
                    if not initial_bom.empty:
                        comparison_report_df = generate_comparison_report(
                            initial_bom, optimized_bom_df, target_qty_optimized)
                        sheets_to_download["Initial vs Optimized"] = comparison_report_df
                        trace_report_df = generate_explosion_trace_report(
                            df, initial_bom, optimized_bom_df, target_qty_optimized)
                        if trace_report_df is not None:
                            sheets_to_download["Explosion Trace"] = trace_report_df
                    excel_data = to_excel_multi_sheet(sheets_to_download)
                    st.download_button(label="📥 Download Full Optimized Report", data=excel_data,
                                       file_name=f"Optimized_BOM_Report_{selected_pf_code}.xlsx")
                else:
                    st.warning("Could not generate an optimized BOM.")

        with tab4:
            st.subheader("BOM Flowchart Visualization")
            st.info(
                "This chart shows the product structure, flowing from left to right.")
            try:
                graph = generate_bom_graph(df, selected_pf_code)
                st.graphviz_chart(graph)
                image_data = graph.pipe(format='png')
                st.download_button(label="📥 Download Flowchart as Image (PNG)", data=image_data,
                                   file_name=f"Flowchart_{selected_pf_code}.png", mime="image/png")
            except Exception as e:
                st.error(
                    f"Could not generate graph. Ensure the Graphviz system package is installed. Error: {e}")
