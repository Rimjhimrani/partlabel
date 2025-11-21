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
    page_title="AgiloSmartTag Studio",
    page_icon="🏷️",
    layout="wide"
)

# --- Style Definitions ---
bold_style_v1 = ParagraphStyle(
    name='Bold_v1', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, leading=16, spaceBefore=5, spaceAfter=2
)
bold_style_v2 = ParagraphStyle(
    name='Bold_v2', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, leading=16, spaceBefore=10, spaceAfter=15,
)
desc_style = ParagraphStyle(
    name='Description', fontName='Helvetica', fontSize=20, alignment=TA_LEFT, leading=16, spaceBefore=2, spaceAfter=2
)
location_header_style = ParagraphStyle(
    name='LocationHeader', fontName='Helvetica', fontSize=16, alignment=TA_CENTER, leading=18
)
location_value_style_v1 = ParagraphStyle(
    name='LocationValue_v1', fontName='Helvetica', fontSize=14, alignment=TA_CENTER, leading=16
)
location_value_style_v2 = ParagraphStyle(
    name='LocationValue_v2', fontName='Helvetica', fontSize=16, alignment=TA_CENTER, leading=18
)


# --- Formatting Functions ---
def format_part_no_v1(part_no):
    if not part_no or not isinstance(part_no, str): part_no = str(part_no)
    if len(part_no) > 5:
        part1, part2 = part_no[:-5], part_no[-5:]
        return Paragraph(f"<b><font size=17>{part1}</font><font size=22>{part2}</font></b>", bold_style_v1)
    return Paragraph(f"<b><font size=17>{part_no}</font></b>", bold_style_v1)

def format_part_no_v2(part_no):
    if not part_no or not isinstance(part_no, str): part_no = str(part_no)
    if part_no.upper() == 'EMPTY':
         return Paragraph(f"<b><font size=34>EMPTY</font></b><br/><br/>", bold_style_v2)
    if len(part_no) > 5:
        part1, part2 = part_no[:-5], part_no[-5:]
        return Paragraph(f"<b><font size=34>{part1}</font><font size=40>{part2}</font></b><br/><br/>", bold_style_v2)
    return Paragraph(f"<b><font size=34>{part_no}</font></b><br/><br/>", bold_style_v2)

def format_description_v1(desc):
    if not desc or not isinstance(desc, str): desc = str(desc)
    font_size = 15 if len(desc) <= 30 else 13 if len(desc) <= 50 else 11 if len(desc) <= 70 else 10 if len(desc) <= 90 else 9
    desc_style_v1 = ParagraphStyle(name='Description_v1', fontName='Helvetica', fontSize=font_size, alignment=TA_LEFT, leading=font_size + 2)
    return Paragraph(desc, desc_style_v1)

def format_description(desc):
    if not desc or not isinstance(desc, str): desc = str(desc)
    return Paragraph(desc, desc_style)

# --- Core Logic Functions ---
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
    if not all([part_no_col, container_col, station_col]):
        st.error("❌ 'Part Number', 'Container Type', or 'Station No' column not found.")
        return None

    df_processed = df.copy()
    rename_dict = {
        part_no_col: 'Part No', desc_col: 'Description',
        model_col: 'Bus Model', station_col: 'Station No', container_col: 'Container'
    }
    df_processed.rename(columns={k: v for k, v in rename_dict.items() if k}, inplace=True)
    df_processed.sort_values(by=['Station No', 'Container'], inplace=True)
    final_df_parts = []
    
    for station_no, station_group in df_processed.groupby('Station No', sort=False):
        if status_text: status_text.text(f"Processing station: {station_no}...")
        rack_idx, level_idx = 0, 0
        sorted_racks = sorted(rack_configs.items())

        for container_type, parts_group in station_group.groupby('Container', sort=True):
            items_to_place = parts_group.to_dict('records')
            while items_to_place:
                slot_found = False
                while rack_idx < len(sorted_racks):
                    current_rack_name, current_config = sorted_racks[rack_idx]
                    allowed_levels = current_config.get('levels', [])
                    capacity_for_this_bin = current_config.get('rack_bin_counts', {}).get(container_type, 0)
                    if capacity_for_this_bin > 0 and level_idx < len(allowed_levels):
                        slot_found = True
                        break
                    rack_idx += 1
                    level_idx = 0
                
                if not slot_found:
                    st.error(f"❌ Ran out of rack space at Station {station_no} for container '{container_type}'. Aborting for this station.")
                    items_to_place = []
                    continue
                
                current_rack_name, current_config = sorted_racks[rack_idx]
                allowed_levels = current_config.get('levels', [])
                level_capacity = current_config.get('rack_bin_counts', {}).get(container_type, 0)
                num_to_place = min(len(items_to_place), level_capacity)
                parts_for_level = items_to_place[:num_to_place]
                items_to_place = items_to_place[num_to_place:]
                level_items = parts_for_level + [{'Part No': 'EMPTY'}] * (level_capacity - len(parts_for_level))

                cell_idx = 1
                for item in level_items:
                    rack_num_val = ''.join(filter(str.isdigit, current_rack_name))
                    rack_num_1st = rack_num_val[0] if len(rack_num_val) > 1 else '0'
                    rack_num_2nd = rack_num_val[1] if len(rack_num_val) > 1 else rack_num_val[0]
                    location_info = {'Rack': base_rack_id, 'Rack No 1st': rack_num_1st, 'Rack No 2nd': rack_num_2nd, 'Level': allowed_levels[level_idx], 'Cell': f"{cell_idx}", 'Station No': station_no}
                    if item['Part No'] == 'EMPTY':
                        item = {**{'Description': '', 'Bus Model': '', 'Container': container_type}, **item, **location_info}
                    else:
                        item.update(location_info)
                    final_df_parts.append(item)
                    cell_idx += 1
                level_idx += 1
    return pd.DataFrame(final_df_parts) if final_df_parts else pd.DataFrame()

# --- CRITICAL FIX: The location key now includes Station No to ensure uniqueness across stations ---
def create_location_key(row):
    return '_'.join([str(row.get(c, '')) for c in ['Station No', 'Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']])

def extract_location_values(row):
    return [str(row.get(c, '')) for c in ['Bus Model', 'Station No', 'Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']]

# --- PDF Generation Functions ---
def generate_labels_from_excel_v1(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
    elements, label_summary = [], {}
    
    df['location_key'] = df.apply(create_location_key, axis=1)
    df.sort_values(by=['Station No', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)

    for i, (location_key, group) in enumerate(df_grouped):
        if progress_bar: progress_bar.progress((i + 1) / total_locations)
        if status_text: status_text.text(f"Processing V1 Label {i+1}/{total_locations}")
        
        part1 = group.iloc[0]
        if str(part1['Part No']).upper() == 'EMPTY': continue

        # --- SUMMARY FIX: Track labels per station and per rack ---
        station_no = str(part1.get('Station No', 'Unknown'))
        rack_num = f"{part1.get('Rack No 1st', '0')}{part1.get('Rack No 2nd', '0')}"
        rack_key = f"Rack {rack_num.zfill(2)}"
        if station_no not in label_summary: label_summary[station_no] = {}
        label_summary[station_no][rack_key] = label_summary[station_no].get(rack_key, 0) + 1

        if len(elements) > 0 and len(elements) % 4 == 0: elements.append(PageBreak())
        
        part2 = group.iloc[1] if len(group) > 1 else part1
        part_table1 = Table([['Part No', format_part_no_v1(str(part1['Part No']))], ['Description', format_description_v1(str(part1['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.3*cm, 0.8*cm])
        part_table2 = Table([['Part No', format_part_no_v1(str(part2['Part No']))], ['Description', format_description_v1(str(part2['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.3*cm, 0.8*cm])
        
        location_data = [[Paragraph('Line Location', location_header_style)] + [Paragraph(str(val), location_value_style_v1) for val in extract_location_values(part1)]]
        location_widths = [4*cm] + [w * (11*cm) / 20.3 for w in [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]]
        location_table = Table(location_data, colWidths=location_widths, rowHeights=0.8*cm)
        
        part_style = TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (0, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, -1), 16)])
        part_table1.setStyle(part_style); part_table2.setStyle(part_style)
        
        loc_style_cmds = [('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]
        for j, color in enumerate([colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]):
            loc_style_cmds.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
        location_table.setStyle(TableStyle(loc_style_cmds))
        
        elements.extend([part_table1, Spacer(1, 0.3*cm), part_table2, Spacer(1, 0.3*cm), location_table, Spacer(1, 0.2*cm)])
        
    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer, label_summary

def generate_labels_from_excel_v2(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
    elements, label_summary = [], {}
    
    df['location_key'] = df.apply(create_location_key, axis=1)
    df.sort_values(by=['Station No', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)

    for i, (location_key, group) in enumerate(df_grouped):
        if progress_bar: progress_bar.progress((i + 1) / total_locations)
        if status_text: status_text.text(f"Processing V2 Label {i+1}/{total_locations}")
        
        part1 = group.iloc[0]
        if str(part1['Part No']).upper() == 'EMPTY': continue
        
        # --- SUMMARY FIX: Track labels per station and per rack ---
        station_no = str(part1.get('Station No', 'Unknown'))
        rack_num = f"{part1.get('Rack No 1st', '0')}{part1.get('Rack No 2nd', '0')}"
        rack_key = f"Rack {rack_num.zfill(2)}"
        if station_no not in label_summary: label_summary[station_no] = {}
        label_summary[station_no][rack_key] = label_summary[station_no].get(rack_key, 0) + 1
            
        if len(elements) > 0 and len(elements) % 4 == 0: elements.append(PageBreak())

        part_table = Table([['Part No', format_part_no_v2(str(part1['Part No']))], ['Description', format_description(str(part1['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.9*cm, 2.1*cm])
        location_data = [[Paragraph('Line Location', location_header_style)] + [Paragraph(str(val), location_value_style_v2) for val in extract_location_values(part1)]]
        location_widths = [4*cm] + [w * (11*cm) / 21 for w in [1.7, 2.9, 1.3, 1.2, 1.3, 1.3, 1.3]]
        location_table = Table(location_data, colWidths=location_widths, rowHeights=0.9*cm)
        
        part_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTER'), ('ALIGN', (1, 1), (1, -1), 'LEFT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 5), ('FONTNAME', (0, 0), (0, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, -1), 16)]))
        loc_style_cmds = [('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]
        for j, color in enumerate([colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]):
            loc_style_cmds.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
        location_table.setStyle(TableStyle(loc_style_cmds))
        
        elements.extend([part_table, Spacer(1, 0.3*cm), location_table, Spacer(1, 0.2*cm)])
        
    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer, label_summary

# --- Main Application UI ---
def main():
    st.title("🏷️ AgiloSmartTag Studio")
    st.markdown("<p style='font-style:italic;'>Designed and Developed by Agilomatrix</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.sidebar.title("📄 Label Options")
    label_type = st.sidebar.selectbox("Choose Label Format:", ["Single Part", "Multiple Parts"])
    base_rack_id = st.sidebar.text_input("Enter Storage Line Side Infrastructure", "R", help="E.g., R for Rack, T for Tray.")
    st.sidebar.caption("EXAMPLE: **R**=RACK, **TR**=TRAY, **SH**=SHELVING")
    
    uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            st.success(f"✅ File loaded! Found {len(df)} rows.")
            
            _, _, _, _, container_col = find_required_columns(df)
            if container_col:
                with st.expander("⚙️ Step 1: Configure Dimensions and Rack Setup", expanded=True):
                    st.subheader("1. Container Dimensions")
                    unique_containers = get_unique_containers(df, container_col)
                    bin_dims = {c: st.text_input(f"Dimensions for {c}", key=f"d_{c}", placeholder="e.g., 300x200x150mm") for c in unique_containers}
                    
                    st.markdown("---"); st.subheader("2. Rack Dimensions & Bin/Level Capacity")
                    st.info("This configuration will be applied independently to each unique station.")
                    num_racks = st.number_input("Number of Racks", min_value=1, value=1, step=1)
                    rack_configs, rack_dims = {}, {}
                    
                    for i in range(num_racks):
                        rack_name = f"Rack {i+1:02d}"
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown(f"**Settings for {rack_name}**")
                            rack_dims[rack_name] = st.text_input(f"Dimensions for {rack_name}", key=f"rd_{rack_name}", placeholder="e.g., 1200x1000x2000mm")
                            levels = st.multiselect(f"Available Levels for {rack_name}", ['A','B','C','D','E','F','G','H'], default=['A','B','C','D','E'], key=f"l_{rack_name}")
                        with col2:
                            st.markdown(f"**Set Total Bin Capacity for {rack_name}**")
                            rack_bin_counts = {c: st.number_input(f"Capacity of '{c}' Bins", 0, key=f"bc_{rack_name}_{c}") for c in unique_containers}
                        rack_configs[rack_name] = {'dimensions': rack_dims[rack_name], 'levels': levels, 'rack_bin_counts': {k:v for k,v in rack_bin_counts.items() if v>0}}
                        st.markdown("---")

                if st.button("🚀 Generate PDF Labels", type="primary"):
                    if any(not d for d in bin_dims.values()) or any(not d for d in rack_dims.values()):
                        st.error("❌ Please provide all dimension information before generating.")
                        st.stop()
                    
                    progress_bar = st.progress(0); status_text = st.empty()
                    try:
                        df_processed = automate_location_assignment(df, base_rack_id, rack_configs, status_text)
                        if df_processed is not None and not df_processed.empty:
                            gen_func = generate_labels_from_excel_v2 if label_type == "Single Part" else generate_labels_from_excel_v1
                            pdf_buffer, label_summary = gen_func(df_processed, progress_bar, status_text)
                            
                            if pdf_buffer.getbuffer().nbytes > 0:
                                # --- SUMMARY FIX: Create a detailed summary DataFrame ---
                                summary_records = []
                                total_labels = 0
                                for station, racks in sorted(label_summary.items()):
                                    for rack, count in sorted(racks.items()):
                                        summary_records.append({'Station': station, 'Rack': rack, 'Number of Labels': count})
                                        total_labels += count
                                
                                status_text.text(f"✅ PDF with {total_labels} labels generated successfully!")
                                file_name = f"{os.path.splitext(uploaded_file.name)[0]}_labels.pdf"
                                st.download_button("📥 Download PDF", pdf_buffer, file_name, "application/pdf")

                                if summary_records:
                                    st.markdown("---"); st.subheader("📊 Generation Summary")
                                    st.markdown(f"A total of **{total_labels}** labels have been generated.")
                                    summary_df = pd.DataFrame(summary_records)
                                    st.table(summary_df)
                            else:
                                st.warning("⚠️ No valid labels could be generated based on the input and configuration. The PDF is empty.")
                        else:
                            st.error("❌ No data was processed. Check your input file and rack capacities.")
                    except Exception as e:
                        st.error(f"❌ An unexpected error occurred: {e}"); st.exception(e)
                    finally:
                        progress_bar.empty(); status_text.empty()
            else:
                st.error("❌ A column containing 'Container' could not be found in the uploaded file.")
        except Exception as e:
            st.error(f"❌ Error reading file: {e}")
    else:
        st.info("👆 Upload a file to begin.")

if __name__ == "__main__":
    main()
