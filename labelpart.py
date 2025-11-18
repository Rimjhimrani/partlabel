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

# --- Style Definitions (No Changes) ---
bold_style_v1 = ParagraphStyle(
    name='Bold_v1', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, leading=14, spaceBefore=2, spaceAfter=2
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

# --- Advanced Core Logic Functions ---
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

# --- THIS IS THE REBUILT CORE LOGIC FUNCTION ---
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
    
    # Process each station independently
    for station_no, station_group in df_processed.groupby('Station No', sort=False):
        if status_text: status_text.text(f"Processing station: {station_no}...")
        
        # 1. Create a master list of all PHYSICAL slots available for this station
        all_physical_slots = []
        for rack_name, config in sorted(rack_configs.items()):
            rack_num_val = ''.join(filter(str.isdigit, rack_name))
            rack_num_1st = rack_num_val[0] if len(rack_num_val) > 1 else '0'
            rack_num_2nd = rack_num_val[1] if len(rack_num_val) > 1 else rack_num_val[0]
            
            for level in sorted(config.get('levels', [])):
                num_cells = config.get('physical_level_capacities', {}).get(level, 0)
                for i in range(num_cells):
                    slot = {
                        'Level': level, 'Cell': f"{i + 1:02d}",
                        'Rack': base_rack_id, 'Rack No 1st': rack_num_1st, 'Rack No 2nd': rack_num_2nd,
                    }
                    all_physical_slots.append(slot)
        
        # 2. Create a list of all REQUIRED bin types based on total rack counts
        all_required_bins = []
        for rack_name, config in sorted(rack_configs.items()):
            for container_type, count in sorted(config.get('rack_bin_counts', {}).items()):
                all_required_bins.extend([container_type] * count)

        # Warn if the number of bins defined exceeds the total physical space
        if len(all_required_bins) > len(all_physical_slots):
            st.warning(f"⚠️ For station {station_no}, you have defined {len(all_required_bins)} total bins, "
                       f"but there is only physical space for {len(all_physical_slots)}. "
                       f"{len(all_required_bins) - len(all_physical_slots)} bins will be ignored.")
        
        # 3. Map the required bin types to the physical slots to create the final slot map
        available_slots_by_container = {}
        num_slots_to_map = min(len(all_physical_slots), len(all_required_bins))
        for i in range(num_slots_to_map):
            physical_slot = all_physical_slots[i]
            container_type = all_required_bins[i]
            
            physical_slot['Container_Type_Slot'] = container_type
            
            if container_type not in available_slots_by_container:
                available_slots_by_container[container_type] = []
            available_slots_by_container[container_type].append(physical_slot)

        # 4. Assign parts from the input file to the available, designated slots
        assigned_physical_slots = set()
        for container_type, parts_group in station_group.groupby('Container', sort=False):
            parts_to_assign = parts_group.to_dict('records')
            slots_for_this_container = available_slots_by_container.get(container_type, [])
            
            if len(parts_to_assign) > len(slots_for_this_container):
                unassigned_count = len(parts_to_assign) - len(slots_for_this_container)
                st.warning(f"⚠️ For station {station_no}, could not place {unassigned_count} parts needing '{container_type}' due to insufficient capacity defined for that bin type.")
                parts_to_assign = parts_to_assign[:len(slots_for_this_container)]

            for i, part in enumerate(parts_to_assign):
                slot = slots_for_this_container[i]
                part.update(slot)
                final_df_parts.append(part)
                # Mark this physical slot as occupied by a part
                slot_key = f"{slot['Rack No 1st']}{slot['Rack No 2nd']}_{slot['Level']}_{slot['Cell']}"
                assigned_physical_slots.add(slot_key)
        
        # 5. Create 'EMPTY' records for all physical slots that did not receive a part
        for slot in all_physical_slots:
            slot_key = f"{slot['Rack No 1st']}{slot['Rack No 2nd']}_{slot['Level']}_{slot['Cell']}"
            if slot_key not in assigned_physical_slots:
                empty_part = {
                    'Part No': 'EMPTY', 'Description': '', 'Bus Model': '', 'Station No': station_no, 'Container': '',
                    **slot
                }
                final_df_parts.append(empty_part)

    if not final_df_parts: return pd.DataFrame()
    return pd.DataFrame(final_df_parts)


def create_location_key(row):
    return '_'.join([str(row.get(c, '')) for c in ['Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']])

def extract_location_values(row):
    return [str(row.get(c, '')) for c in ['Bus Model', 'Station No', 'Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']]

# --- PDF Generation Functions (No Changes) ---
def generate_labels_from_excel_v1(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
    elements = []
    
    df['location_key'] = df.apply(create_location_key, axis=1)
    df.sort_values(by=['Station No', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)
    label_count = 0
    label_summary = {}

    for i, (location_key, group) in enumerate(df_grouped):
        if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
        if status_text: status_text.text(f"Processing V1 Label {i+1}/{total_locations}")
        
        part1 = group.iloc[0]
        if str(part1['Part No']).upper() == 'EMPTY':
            continue

        rack_num = f"{part1.get('Rack No 1st', '0')}{part1.get('Rack No 2nd', '0')}"
        rack_key = f"Rack {rack_num.zfill(2)}"
        label_summary[rack_key] = label_summary.get(rack_key, 0) + 1

        if label_count > 0 and label_count % 4 == 0: elements.append(PageBreak())
        
        part2 = group.iloc[1] if len(group) > 1 else part1
        
        part_table1 = Table([['Part No', format_part_no_v1(str(part1['Part No']))], ['Description', format_description_v1(str(part1['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.3*cm, 0.8*cm])
        part_table2 = Table([['Part No', format_part_no_v1(str(part2['Part No']))], ['Description', format_description_v1(str(part2['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.3*cm, 0.8*cm])
        
        location_values = extract_location_values(part1)
        location_data = [['Line Location'] + location_values]
        col_proportions = [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]
        location_widths = [4 * cm] + [w * (11 * cm) / sum(col_proportions) for w in col_proportions]
        location_table = Table(location_data, colWidths=location_widths, rowHeights=0.8*cm)
        
        part_style = TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (0, -1), 16)])
        part_table1.setStyle(part_style)
        part_table2.setStyle(part_style)
        
        location_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
        location_style = [('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (0, 0), 16), ('FONTSIZE', (1, 0), (-1, -1), 14)]
        for j, color in enumerate(location_colors): location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
        location_table.setStyle(TableStyle(location_style))
        
        elements.append(part_table1)
        elements.append(Spacer(1, 0.3 * cm))
        elements.append(part_table2)
        elements.append(Spacer(1, 0.3 * cm))
        elements.append(location_table)
        elements.append(Spacer(1, 0.2 * cm))
        label_count += 1
        
    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer, label_summary

def generate_labels_from_excel_v2(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
    elements = []
    
    df['location_key'] = df.apply(create_location_key, axis=1)
    df.sort_values(by=['Station No', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)
    label_count = 0
    label_summary = {}

    for i, (location_key, group) in enumerate(df_grouped):
        if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
        if status_text: status_text.text(f"Processing V2 Label {i+1}/{total_locations}")

        part1 = group.iloc[0]
        if str(part1['Part No']).upper() == 'EMPTY':
            continue
        
        rack_num = f"{part1.get('Rack No 1st', '0')}{part1.get('Rack No 2nd', '0')}"
        rack_key = f"Rack {rack_num.zfill(2)}"
        label_summary[rack_key] = label_summary.get(rack_key, 0) + 1
            
        if label_count > 0 and label_count % 4 == 0: elements.append(PageBreak())

        part_table = Table([['Part No', format_part_no_v2(str(part1['Part No']))], ['Description', format_description(str(part1['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.9*cm, 2.1*cm])
        
        location_values = extract_location_values(part1)
        location_data = [['Line Location'] + location_values]
        col_widths = [1.7, 2.9, 1.3, 1.2, 1.3, 1.3, 1.3]
        location_widths = [4 * cm] + [w * (11 * cm) / sum(col_widths) for w in col_widths]
        location_table = Table(location_data, colWidths=location_widths, rowHeights=0.9*cm)
        
        part_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTER'), ('ALIGN', (1, 1), (1, -1), 'LEFT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 5), ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (0, -1), 16)]))
        
        location_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
        location_style = [('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (0, 0), 16), ('FONTSIZE', (1, 0), (-1, -1), 16)]
        for j, color in enumerate(location_colors): location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
        location_table.setStyle(TableStyle(location_style))
        
        elements.append(part_table)
        elements.append(Spacer(1, 0.3 * cm))
        elements.append(location_table)
        elements.append(Spacer(1, 0.2 * cm))
        label_count += 1
        
    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer, label_summary

# --- Main Application UI (Updated) ---
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
                
                st.sidebar.markdown("---")
                st.sidebar.subheader("Container Dimensions")
                st.sidebar.info("Provide dimensions for each container type found in your file.")
                for container in unique_containers:
                    st.sidebar.text_input(f"Dimensions for {container}", key=f"bindim_{container}", placeholder="e.g., 300x200x150mm")
                
                st.sidebar.markdown("---")
                st.sidebar.subheader("Rack & Level Configuration")
                st.sidebar.info("This configuration will be applied independently to each unique station.")
                
                num_racks = st.sidebar.number_input("Number of Racks", min_value=1, value=1, step=1)

                rack_configs = {}

                for i in range(num_racks):
                    rack_name = f"Rack {i+1:02d}"
                    with st.sidebar.expander(f"Settings for {rack_name}", expanded=i==0):
                        st.text_input(f"Dimensions for {rack_name}", key=f"dim_{rack_name}", placeholder="e.g., 1200x1000x2000mm")
                        
                        levels = st.multiselect(f"Levels for {rack_name}", options=['A','B','C','D','E','F','G','H'], default=['A','B','C','D'], key=f"lvl_{rack_name}")
                        
                        # --- THIS IS THE NEW UI LOGIC ---
                        physical_level_capacities = {}
                        if levels:
                            st.markdown("---")
                            st.write("**1. Set Physical Cells per Level**")
                            for level in levels:
                                p_capacity = st.number_input(f"Total Cells on Level {level}", min_value=1, value=10, step=1, key=f"pcap_{rack_name}_{level}")
                                physical_level_capacities[level] = p_capacity

                        rack_bin_counts = {}
                        if levels and unique_containers:
                            st.markdown("---")
                            st.write("**2. Set Total Bin Count for this Rack**")
                            for container in unique_containers:
                                b_count = st.number_input(f"Total Count of '{container}'", min_value=0, value=0, step=1, key=f"bcount_{rack_name}_{container}")
                                if b_count > 0:
                                    rack_bin_counts[container] = b_count
                        
                        rack_configs[rack_name] = {
                            'levels': levels, 
                            'physical_level_capacities': physical_level_capacities,
                            'rack_bin_counts': rack_bin_counts
                        }

                if st.button("🚀 Generate PDF Labels", type="primary"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    try:
                        df_processed = automate_location_assignment(df, base_rack_id, rack_configs, status_text)
                        
                        if df_processed is not None and not df_processed.empty:
                            gen_func = generate_labels_from_excel_v2 if label_type == "Single Part" else generate_labels_from_excel_v1
                            pdf_buffer, label_summary = gen_func(df_processed, progress_bar, status_text)
                            
                            if pdf_buffer:
                                total_labels = sum(label_summary.values())
                                status_text.text(f"✅ PDF with {total_labels} labels generated successfully!")
                                file_name = f"{os.path.splitext(uploaded_file.name)[0]}_{label_type.lower().replace(' ','_')}_labels.pdf"
                                st.download_button(label="📥 Download PDF", data=pdf_buffer.getvalue(), file_name=file_name, mime="application/pdf")

                                if total_labels > 0:
                                    st.markdown("---")
                                    st.subheader("📊 Generation Summary")
                                    st.markdown(f"A total of **{total_labels}** labels have been generated. Here is the breakdown by rack:")
                                    summary_df = pd.DataFrame(list(label_summary.items()), columns=['Rack', 'Number of Labels'])
                                    summary_df = summary_df.sort_values(by='Rack').reset_index(drop=True)
                                    st.table(summary_df)
                        else:
                            st.error("❌ No data was processed. Check your input file and ensure rack capacities are configured.")
                    except Exception as e:
                        st.error(f"❌ An unexpected error occurred: {e}")
                        st.exception(e)
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
