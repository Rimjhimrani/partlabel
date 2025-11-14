import streamlit as st
import pandas as pd
import os
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT

# --- Page Configuration (No Changes) ---
st.set_page_config(
    page_title="Part Label Generator",
    page_icon="🏷️",
    layout="wide"
)

# --- Style Definitions (No Changes) ---
bold_style_v1 = ParagraphStyle(
    name='Bold_v1', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, leading=20, spaceBefore=2, spaceAfter=2
)
bold_style_v2 = ParagraphStyle(
    name='Bold_v2', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, leading=12, spaceBefore=0, spaceAfter=15,
)
desc_style = ParagraphStyle(
    name='Description', fontName='Helvetica', fontSize=20, alignment=TA_LEFT, leading=16, spaceBefore=2, spaceAfter=2
)

# --- Formatting Functions (No Changes) ---
def format_part_no_v1(part_no):
    if not part_no or not isinstance(part_no, str): part_no = str(part_no)
    if len(part_no) > 5:
        part1, part2 = part_no[:-5], part_no[-5:]
        return Paragraph(f"<b><font size=17>{part1}</font><font size=22>{part2}</font></b>", bold_style_v1)
    return Paragraph(f"<b><font size=17>{part_no}</font></b>", bold_style_v1)

def format_part_no_v2(part_no):
    if not part_no or not isinstance(part_no, str): part_no = str(part_no)
    if len(part_no) > 5:
        part1, part2 = part_no[:-5], part_no[-5:]
        return Paragraph(f"<b><font size=34>{part1}</font><font size=40>{part2}</font></b><br/><br/>", bold_style_v2)
    return Paragraph(f"<b><font size=34>{part_no}</font></b><br/><br/>", bold_style_v2)

def format_description_v1(desc):
    if not desc or not isinstance(desc, str): desc = str(desc)
    desc_len = len(desc)
    if desc_len <= 30: font_size = 15
    elif desc_len <= 50: font_size = 13
    elif desc_len <= 70: font_size = 11
    elif desc_len <= 90: font_size = 10
    else: font_size = 9
    style = ParagraphStyle(name='Desc_v1', fontName='Helvetica', fontSize=font_size, alignment=TA_LEFT, leading=font_size + 2)
    return Paragraph(desc, style)

def format_description(desc):
    if not desc or not isinstance(desc, str): desc = str(desc)
    return Paragraph(desc, desc_style)

# --- Core Logic Functions (UPDATED) ---

def find_required_columns(df):
    cols = {col.upper().strip(): col for col in df.columns}
    part_no_key = next((k for k in cols if 'PART' in k and ('NO' in k or 'NUM' in k)), None)
    desc_key = next((k for k in cols if 'DESC' in k), None)
    bus_model_key = next((k for k in cols if 'BUS' in k and 'MODEL' in k), None)
    station_no_key = next((k for k in cols if 'STATION' in k), None)
    container_type_key = next((k for k in cols if 'CONTAINER' in k), None)
    return (cols.get(part_no_key), cols.get(desc_key), cols.get(bus_model_key),
            cols.get(station_no_key), cols.get(container_type_key))

def get_unique_containers(df, container_col):
    if not container_col or container_col not in df.columns: return []
    unique_containers = df[container_col].dropna().astype(str).unique()
    return sorted([c for c in unique_containers])

def automate_location_assignment_v2(df, base_rack_id, rack_configs, status_text=None):
    part_no_col, desc_col, model_col, station_col, container_col = find_required_columns(df)
    if not part_no_col or not container_col:
        st.error("❌ 'Part Number' or 'Container Type' column not found.")
        return None

    df_processed = df.copy()
    
    # Standardize column names for processing
    rename_dict = {
        part_no_col: 'Part No', desc_col: 'Description',
        model_col: 'Bus Model', station_col: 'Station No', container_col: 'Container'
    }
    df_processed.rename(columns={k: v for k, v in rename_dict.items() if k}, inplace=True)

    # Initialize state trackers for each rack
    rack_fill_status = {
        rack_name: {'level_idx': 0, 'cell_count': 0}
        for rack_name in rack_configs
    }
    
    all_assigned_parts = []
    
    # Group parts by container type to process them together
    for container_type, group in df_processed.groupby('Container'):
        parts_to_assign = group.to_dict('records')
        
        # Find racks that can hold this container type
        eligible_racks = [
            rack_name for rack_name, config in rack_configs.items()
            if container_type in config['capacities'] and config['capacities'][container_type] > 0
        ]
        
        if not eligible_racks:
            if status_text: st.warning(f"No rack configured for container type: {container_type}. Skipping.")
            continue
            
        current_rack_idx = 0
        
        for part in parts_to_assign:
            assigned = False
            while not assigned and current_rack_idx < len(eligible_racks):
                rack_name = eligible_racks[current_rack_idx]
                config = rack_configs[rack_name]
                status = rack_fill_status[rack_name]
                
                capacity = config['capacities'].get(container_type, 0)
                levels = config['levels']
                
                if not levels: # No levels selected for this rack
                    current_rack_idx += 1
                    continue
                
                if status['level_idx'] < len(levels):
                    # Assign part to the current level and cell
                    part['Rack'] = base_rack_id
                    rack_num_str = rack_name.replace("Rack ", "")
                    part['Rack No 1st'] = rack_num_str[0] if len(rack_num_str) > 1 else '0'
                    part['Rack No 2nd'] = rack_num_str[1] if len(rack_num_str) > 1 else rack_num_str[0]
                    part['Level'] = levels[status['level_idx']]
                    part['Cell'] = f"{(status['cell_count'] % capacity) + 1:02d}"
                    all_assigned_parts.append(part)
                    
                    status['cell_count'] += 1
                    assigned = True
                    
                    # Move to the next level if the current one is full
                    if status['cell_count'] >= capacity:
                        status['cell_count'] = 0
                        status['level_idx'] += 1
                else:
                    # All levels in this rack are full for this container type, move to the next rack
                    current_rack_idx += 1
    
    final_df = pd.DataFrame(all_assigned_parts)
    
    # --- Generate Blanks ---
    if status_text: status_text.text("Generating blank locations...")
    blank_rows = []
    for rack_name, config in rack_configs.items():
        for container_type, capacity in config['capacities'].items():
            if capacity == 0: continue
            for level in config['levels']:
                # Find how many parts were actually placed in this location
                existing_parts = final_df[
                    (final_df['Rack No 2nd'] == rack_name[-1]) &
                    (final_df['Level'] == level) &
                    (final_df['Container'] == container_type)
                ]
                num_existing = len(existing_parts)
                
                # Add blank rows if capacity is not met
                for i in range(num_existing, capacity):
                    rack_num_str = rack_name.replace("Rack ", "")
                    blank_row = {
                        'Part No': 'EMPTY', 'Description': '', 'Bus Model': '', 'Station No': '',
                        'Rack': base_rack_id,
                        'Rack No 1st': rack_num_str[0] if len(rack_num_str) > 1 else '0',
                        'Rack No 2nd': rack_num_str[1] if len(rack_num_str) > 1 else rack_num_str[0],
                        'Level': level,
                        'Cell': f"{i + 1:02d}"
                    }
                    blank_rows.append(blank_row)
    
    if blank_rows:
        final_df = pd.concat([final_df, pd.DataFrame(blank_rows)], ignore_index=True)
        
    return final_df

# --- PDF Generation Functions (No Major Changes) ---
def create_location_key(row):
    return '_'.join([str(row.get(c, '')) for c in ['Bus Model', 'Station No', 'Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']])

def extract_location_values(row):
    return [str(row.get(c, '')) for c in ['Bus Model', 'Station No', 'Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']]

def generate_labels_from_excel_v1(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    
    df['location_key'] = df.apply(create_location_key, axis=1)
    df.sort_values(by=['Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)
    label_count = 0

    for i, (location_key, group) in enumerate(df_grouped):
        if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
        if status_text: status_text.text(f"Processing location {i+1}/{total_locations}")
        
        if label_count > 0 and label_count % 4 == 0: elements.append(PageBreak())
        label_count += 1
        
        part1 = group.iloc[0]
        part2 = group.iloc[1] if len(group) > 1 else part1
        location_values = extract_location_values(part1)

        p1_no = format_part_no_v1(str(part1['Part No']))
        p1_desc = format_description_v1(str(part1['Description']))
        p2_no = format_part_no_v1(str(part2['Part No'])) if part1['Part No'] != 'EMPTY' else format_part_no_v1('')
        p2_desc = format_description_v1(str(part2['Description']))
        
        part_table1 = Table([['Part No', p1_no], ['Description', p1_desc]], colWidths=[4*cm, 11*cm], rowHeights=[1.3*cm, 0.8*cm])
        part_table2 = Table([['Part No', p2_no], ['Description', p2_desc]], colWidths=[4*cm, 11*cm], rowHeights=[1.3*cm, 0.8*cm])
        location_table = Table([['Line Location'] + location_values], colWidths=[4*cm, 2.3*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.2*cm], rowHeights=0.8*cm)
        
        # Styling (omitted for brevity, same as original)
        part_table1.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        part_table2.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        location_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))

        elements.extend([part_table1, Spacer(1, 0.2*cm), part_table2, location_table, Spacer(1, 0.5*cm)])
    
    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer

def generate_labels_from_excel_v2(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    
    df['location_key'] = df.apply(create_location_key, axis=1)
    df.sort_values(by=['Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)
    label_count = 0

    for i, (location_key, group) in enumerate(df_grouped):
        if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
        if status_text: status_text.text(f"Processing location {i+1}/{total_locations}")

        if label_count > 0 and label_count % 4 == 0: elements.append(PageBreak())
        label_count += 1
        
        part1 = group.iloc[0]
        location_values = extract_location_values(part1)

        p1_no = format_part_no_v2(str(part1['Part No']))
        p1_desc = format_description(str(part1['Description']))

        part_table = Table([['Part No', p1_no], ['Description', p1_desc]], colWidths=[4*cm, 11*cm], rowHeights=[1.9*cm, 2.1*cm])
        location_table = Table([['Line Location'] + location_values], colWidths=[4*cm, 2.3*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.2*cm], rowHeights=0.9*cm)
        
        # Styling (omitted for brevity, same as original)
        part_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        location_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        
        elements.extend([part_table, location_table, Spacer(1, 0.5 * cm)])

    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer

# --- Main Application UI ---
def main():
    st.title("🏷️ Rack Label Generator")
    st.markdown("<p style='font-style:italic;'>Designed by Agilomatrix</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.sidebar.title("📄 Label Options")
    label_type = st.sidebar.selectbox("Choose Label Format:", ["Single Part", "Multiple Parts"])
    
    st.sidebar.title("⚙️ Automation Settings")
    base_rack_id = st.sidebar.text_input("Enter Storage Line Side Infrastructure", "R")
    
    uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            st.success(f"✅ File loaded! Found {len(df)} rows.")
            
            _, _, _, _, container_col = find_required_columns(df)
            
            if container_col:
                unique_containers = get_unique_containers(df, container_col)
                num_racks = len(unique_containers)
                
                st.sidebar.markdown("---")
                st.sidebar.subheader("Rack & Bin Configuration")
                st.sidebar.info(f"Found {len(unique_containers)} container types. They will be assigned to {num_racks} racks.")

                rack_configs = {}
                for i, container in enumerate(unique_containers):
                    rack_name = f"Rack {i+1:02d}"
                    st.sidebar.markdown(f"#### Settings for {rack_name}")
                    
                    rack_dim = st.sidebar.text_input(f"Dimensions for {rack_name}", key=f"dim_{rack_name}")
                    
                    capacities = {}
                    for bin_type in unique_containers:
                         capacities[bin_type] = st.sidebar.number_input(
                             f"Capacity of '{bin_type}' in {rack_name}", 
                             min_value=0, value=1 if bin_type == container else 0, step=1, 
                             key=f"cap_{rack_name}_{bin_type}"
                         )
                    
                    levels = st.sidebar.multiselect(
                        f"Levels to use for {rack_name}",
                        options=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
                        default=['A', 'B', 'C', 'D'],
                        key=f"lvl_{rack_name}"
                    )
                    
                    rack_configs[rack_name] = {'dimensions': rack_dim, 'capacities': capacities, 'levels': levels}
                    st.sidebar.markdown("---")

                if st.button("🚀 Generate PDF Labels", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    try:
                        df_processed = automate_location_assignment_v2(df, base_rack_id, rack_configs, status_text)
                        
                        if df_processed is not None and not df_processed.empty:
                            if label_type == "Single Part":
                                pdf_buffer = generate_labels_from_excel_v2(df_processed, progress_bar, status_text)
                                filename = f"{os.path.splitext(uploaded_file.name)[0]}_single_part_labels.pdf"
                            else:
                                pdf_buffer = generate_labels_from_excel_v1(df_processed, progress_bar, status_text)
                                filename = f"{os.path.splitext(uploaded_file.name)[0]}_multi_part_labels.pdf"
                            
                            if pdf_buffer:
                                status_text.text("✅ PDF generated successfully!")
                                st.download_button(label="📥 Download PDF", data=pdf_buffer.getvalue(), file_name=filename, mime="application/pdf")
                        else:
                            st.error("❌ No data was processed. Check your rack configurations and capacities.")
                            
                    except Exception as e:
                        st.error(f"❌ An unexpected error occurred: {e}")
                    finally:
                        progress_bar.empty()
                        status_text.empty()
            else:
                st.error("Could not find a 'Container Type' column in the file.")
        except Exception as e:
            st.error(f"❌ Error reading file: {e}")
    else:
        st.info("👆 Upload a file to begin.")

if __name__ == "__main__":
    main()
