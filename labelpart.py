import streamlit as st
import pandas as pd
import os
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# --- Page Configuration ---
st.set_page_config(
    page_title="Part Label Generator",
    page_icon="🏷️",
    layout="wide"
)

# --- Style Definitions (UPDATED for better visuals) ---
bold_style_v1 = ParagraphStyle(
    name='Bold_v1', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, leading=20, spaceBefore=2, spaceAfter=2
)

# Style for the large part number in the V2 label
part_no_style_v2 = ParagraphStyle(
    name='PartNo_v2',
    fontName='Helvetica-Bold',
    fontSize=40,
    alignment=TA_LEFT,
    leading=40,  # Keep leading tight
    spaceBefore=0,
    spaceAfter=0,
    leftIndent=5
)

# Style for the description text in the V2 label
desc_style = ParagraphStyle(
    name='Description',
    fontName='Helvetica',
    fontSize=18,  # Optimized font size
    alignment=TA_LEFT,
    leading=22,  # More space for wrapped lines
    spaceBefore=2,
    spaceAfter=2
)

# --- Formatting Functions (UPDATED for new styles) ---
def format_part_no_v1(part_no):
    """Formats part number for the multi-part label."""
    if not part_no or not isinstance(part_no, str): part_no = str(part_no)
    if len(part_no) > 5:
        part1, part2 = part_no[:-5], part_no[-5:]
        return Paragraph(f"<b><font size=17>{part1}</font><font size=22>{part2}</font></b>", bold_style_v1)
    return Paragraph(f"<b><font size=17>{part_no}</font></b>", bold_style_v1)

def format_part_no_v2(part_no):
    """Formats part number for the single-part label with a very large font."""
    if not part_no or not isinstance(part_no, str): part_no = str(part_no)
    if part_no.upper() == 'EMPTY':
        return Paragraph("<b><font size=40>EMPTY</font></b>", part_no_style_v2)
    if len(part_no) > 5:
        part1, part2 = part_no[:-5], part_no[-5:]
        return Paragraph(f"<b><font size=40>{part1}</font><font size=52>{part2}</font></b>", part_no_style_v2)
    return Paragraph(f"<b><font size=40>{part_no}</font></b>", part_no_style_v2)

def format_description_v1(desc):
    """Formats description with dynamic font sizing for multi-part label."""
    if not desc or not isinstance(desc, str): desc = str(desc)
    font_size = 15 if len(desc) <= 30 else 13 if len(desc) <= 50 else 11 if len(desc) <= 70 else 9
    style = ParagraphStyle(name='Desc_v1', fontName='Helvetica', fontSize=font_size, alignment=TA_LEFT, leading=font_size + 2)
    return Paragraph(desc, style)

def format_description(desc):
    """Formats description text for single-part label."""
    if not desc or not isinstance(desc, str): desc = str(desc)
    return Paragraph(desc, desc_style)

# --- Core Logic Functions (No Changes Here) ---
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
    return sorted(df[container_col].dropna().astype(str).unique())

def automate_location_assignment(df, base_rack_id, rack_configs, status_text=None):
    part_no_col, desc_col, model_col, station_col, container_col = find_required_columns(df)
    if not part_no_col or not container_col:
        st.error("❌ 'Part Number' or 'Container Type' column not found.")
        return None

    df_processed = df.copy()
    rename_dict = {
        part_no_col: 'Part No', desc_col: 'Description',
        model_col: 'Bus Model', station_col: 'Station No', container_col: 'Container'
    }
    df_processed.rename(columns={k: v for k, v in rename_dict.items() if k}, inplace=True)

    rack_fill_status = {name: {'level_idx': 0, 'cell_count': 0} for name in rack_configs}
    all_assigned_parts = []
    
    for container_type, group in df_processed.groupby('Container'):
        parts_to_assign = group.to_dict('records')
        eligible_racks = [
            name for name, config in rack_configs.items() if config['capacities'].get(container_type, 0) > 0
        ]
        if not eligible_racks: continue
            
        current_rack_idx = 0
        for part in parts_to_assign:
            assigned = False
            while not assigned and current_rack_idx < len(eligible_racks):
                rack_name = eligible_racks[current_rack_idx]
                config, status = rack_configs[rack_name], rack_fill_status[rack_name]
                capacity, levels = config['capacities'].get(container_type, 0), config['levels']
                
                if not levels or status['level_idx'] >= len(levels):
                    current_rack_idx += 1
                    continue

                rack_num_str = ''.join(filter(str.isdigit, rack_name))
                part.update({
                    'Rack': base_rack_id,
                    'Rack No 1st': rack_num_str[0] if len(rack_num_str) > 1 else '0',
                    'Rack No 2nd': rack_num_str[1] if len(rack_num_str) > 1 else rack_num_str[0],
                    'Level': levels[status['level_idx']],
                    'Cell': f"{(status['cell_count'] % capacity) + 1:02d}"
                })
                all_assigned_parts.append(part)
                
                status['cell_count'] += 1
                assigned = True
                
                if status['cell_count'] >= capacity:
                    status['cell_count'], status['level_idx'] = 0, status['level_idx'] + 1
    
    final_df = pd.DataFrame(all_assigned_parts)
    if status_text: status_text.text("Generating blank locations...")
    blank_rows = []
    for rack_name, config in rack_configs.items():
        for container_type, capacity in config['capacities'].items():
            if capacity == 0: continue
            for level in config['levels']:
                existing_parts_mask = (final_df['Rack No 2nd'] == ''.join(filter(str.isdigit, rack_name))[-1]) & (final_df['Level'] == level) & (final_df['Container'] == container_type)
                num_existing = len(final_df[existing_parts_mask])
                for i in range(num_existing, capacity):
                    rack_num_str = ''.join(filter(str.isdigit, rack_name))
                    blank_rows.append({
                        'Part No': 'EMPTY', 'Description': '', 'Bus Model': '', 'Station No': '', 'Container': container_type,
                        'Rack': base_rack_id,
                        'Rack No 1st': rack_num_str[0] if len(rack_num_str) > 1 else '0',
                        'Rack No 2nd': rack_num_str[1] if len(rack_num_str) > 1 else rack_num_str[0],
                        'Level': level, 'Cell': f"{i + 1:02d}"
                    })
    
    if blank_rows:
        final_df = pd.concat([final_df, pd.DataFrame(blank_rows)], ignore_index=True)
    return final_df

# --- PDF Generation Functions (UPDATED with new styles) ---
def create_location_key(row):
    return '_'.join([str(row.get(c, '')) for c in ['Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']])

def extract_location_values(row):
    return [str(row.get(c, '')) for c in ['Bus Model', 'Station No', 'Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']]

def generate_labels_from_excel_v1(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
    elements = []
    
    df['location_key'] = df.apply(create_location_key, axis=1)
    df.sort_values(by=['Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    # Group by location, but take up to 2 parts for that location
    df_grouped = df.groupby('location_key').apply(lambda x: x.head(2)).reset_index(drop=True).groupby('location_key')
    
    total_locations, label_count = len(df_grouped), 0

    for i, (location_key, group) in enumerate(df_grouped):
        if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
        if status_text: status_text.text(f"Processing V1 Label {i+1}/{total_locations}")
        if label_count > 0 and label_count % 4 == 0: elements.append(PageBreak())
        
        part1 = group.iloc[0]
        part2 = group.iloc[1] if len(group) > 1 else part1
        
        part_table1 = Table([['Part No', format_part_no_v1(str(part1['Part No']))], ['Description', format_description_v1(str(part1['Description']))]], colWidths=[3*cm, 14*cm], rowHeights=[1.3*cm, 0.8*cm])
        part_table2 = Table([['Part No', format_part_no_v1(str(part2['Part No']))], ['Description', format_description_v1(str(part2['Description']))]], colWidths=[3*cm, 14*cm], rowHeights=[1.3*cm, 0.8*cm])
        
        loc_values = extract_location_values(part1)
        location_table = Table([['Line Location'] + loc_values], colWidths=[3*cm, 2.3*cm, 1.9*cm, 1.9*cm, 1.9*cm, 1.9*cm, 1.9*cm, 1.9*cm], rowHeights=1.2*cm)

        # Apply styles
        part_style = TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (0,-1), 'CENTER'), ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (0,-1), 14)])
        part_table1.setStyle(part_style)
        part_table2.setStyle(part_style)
        
        location_style = [('GRID', (0,0), (-1,-1), 1, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (-1,-1), 16), ('TEXTCOLOR', (0,0), (-1,-1), colors.black)]
        location_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
        for j, color in enumerate(location_colors): location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
        location_table.setStyle(TableStyle(location_style))

        elements.extend([part_table1, part_table2, location_table, Spacer(1, 0.5 * cm)])
        label_count += 1

    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer

def generate_labels_from_excel_v2(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
    elements = []
    
    df['location_key'] = df.apply(create_location_key, axis=1)
    df.sort_values(by=['Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    df_grouped = df.groupby('location_key')
    total_locations, label_count = len(df_grouped), 0
    
    for i, (location_key, group) in enumerate(df_grouped):
        if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
        if status_text: status_text.text(f"Processing V2 Label {i+1}/{total_locations}")
        if label_count > 0 and label_count % 4 == 0: elements.append(PageBreak())

        part1 = group.iloc[0]
        part_table = Table([['Part No', format_part_no_v2(str(part1['Part No']))], ['Description', format_description(str(part1['Description']))]], colWidths=[3*cm, 14*cm], rowHeights=[2.5*cm, 2.0*cm])
        
        loc_values = extract_location_values(part1)
        location_table = Table([['Line Location'] + loc_values], colWidths=[3*cm, 2.3*cm, 1.9*cm, 1.9*cm, 1.9*cm, 1.9*cm, 1.9*cm, 1.9*cm], rowHeights=1.2*cm)
        
        part_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (0,-1), 'CENTER'), ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (0,-1), 14), ('LEFTPADDING', (1,0), (1,-1), 10)]))

        location_style = [('GRID', (0,0), (-1,-1), 1, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (0,0), 16), ('FONTSIZE', (1,0), (-1,-1), 18), ('TEXTCOLOR', (0,0), (-1,-1), colors.black)]
        location_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
        for j, color in enumerate(location_colors): location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
        location_table.setStyle(TableStyle(location_style))

        elements.extend([part_table, location_table, Spacer(1, 0.5 * cm)])
        label_count += 1
        
    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer

# --- Main Application UI (UPDATED) ---
def main():
    st.title("🏷️ Rack Label Generator")
    st.markdown("<p style='font-style:italic;'>Designed by Agilomatrix</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.sidebar.title("📄 Label Options")
    label_type = st.sidebar.selectbox("Choose Label Format:", ["Single Part", "Multiple Parts"])
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
                
                # --- NEW: Ask for container dimensions ---
                st.sidebar.markdown("---")
                st.sidebar.subheader("Container Dimensions")
                for container in unique_containers:
                    st.sidebar.text_input(f"Dimensions for {container}", key=f"bindim_{container}", placeholder="e.g., 300x200x150mm")
                
                st.sidebar.markdown("---")
                st.sidebar.subheader("Rack & Bin Configuration")
                st.sidebar.info(f"Found {len(unique_containers)} container types. They will be assigned to {num_racks} racks.")

                rack_configs = {}
                for i, container in enumerate(unique_containers):
                    rack_name = f"Rack {i+1:02d}"
                    with st.sidebar.expander(f"Settings for {rack_name}", expanded=i==0):
                        rack_dim = st.text_input(f"Dimensions for {rack_name}", key=f"dim_{rack_name}")
                        capacities = {bin_type: st.number_input(f"Capacity of '{bin_type}' in {rack_name}", min_value=0, value=1 if bin_type == container else 0, step=1, key=f"cap_{rack_name}_{bin_type}") for bin_type in unique_containers}
                        levels = st.multiselect(f"Levels for {rack_name}", options=['A','B','C','D','E','F','G','H'], default=['A','B','C','D'], key=f"lvl_{rack_name}")
                        rack_configs[rack_name] = {'dimensions': rack_dim, 'capacities': capacities, 'levels': levels}

                if st.button("🚀 Generate PDF Labels", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    try:
                        df_processed = automate_location_assignment(df, base_rack_id, rack_configs, status_text)
                        if df_processed is not None and not df_processed.empty:
                            gen_func = generate_labels_from_excel_v2 if label_type == "Single Part" else generate_labels_from_excel_v1
                            pdf_buffer = gen_func(df_processed, progress_bar, status_text)
                            if pdf_buffer:
                                status_text.text("✅ PDF generated successfully!")
                                file_name = f"{os.path.splitext(uploaded_file.name)[0]}_{label_type.lower().replace(' ','_')}_labels.pdf"
                                st.download_button(label="📥 Download PDF", data=pdf_buffer.getvalue(), file_name=file_name, mime="application/pdf")
                        else:
                            st.error("❌ No data was processed. Check rack capacities and level selections.")
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
