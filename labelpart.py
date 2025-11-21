import streamlit as st
import pandas as pd
import os
import io
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak, Image
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from io import BytesIO

# --- Dependency Check for Bin Labels ---
try:
    import qrcode
    from PIL import Image as PILImage
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False


# --- Page Configuration ---
st.set_page_config(
    page_title="AgiloSmartTag Studio",
    page_icon="🏷️",
    layout="wide"
)

# --- Style Definitions (Shared & Rack-Specific) ---
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

# --- Style Definitions (Bin-Label Specific) ---
bin_bold_style = ParagraphStyle(name='Bold', fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, leading=14)
bin_desc_style = ParagraphStyle(name='Description', fontName='Helvetica', fontSize=11, alignment=TA_CENTER, leading=12)
bin_qty_style = ParagraphStyle(name='Quantity', fontName='Helvetica', fontSize=11, alignment=TA_CENTER, leading=12)


# --- Formatting Functions (Rack Labels) ---
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


# --- Core Logic Functions (Shared) ---
def find_required_columns(df):
    # Returns a dictionary of original column names found in the dataframe
    cols_map = {col.strip().upper(): col for col in df.columns}
    
    def find_col(patterns):
        for p in patterns:
            if p in cols_map:
                return cols_map[p]
        return None

    part_no_col = find_col([k for k in cols_map if 'PART' in k and ('NO' in k or 'NUM' in k)])
    desc_col = find_col([k for k in cols_map if 'DESC' in k])
    bus_model_col = find_col([k for k in cols_map if 'BUS' in k and 'MODEL' in k])
    station_no_col = find_col([k for k in cols_map if 'STATION' in k])
    container_col = find_col([k for k in cols_map if 'CONTAINER' in k])
    qty_bin_col = find_col([k for k in cols_map if 'QTY/BIN' in k or 'QTY_BIN' in k or ('QTY' in k and 'BIN' in k)])
    qty_veh_col = find_col([k for k in cols_map if 'QTY/VEH' in k or 'QTY_VEH' in k or ('QTY' in k and 'VEH' in k)])

    return {
        'Part No': part_no_col, 'Description': desc_col, 'Bus Model': bus_model_col,
        'Station No': station_no_col, 'Container': container_col, 'Qty/Bin': qty_bin_col,
        'Qty/Veh': qty_veh_col
    }

def get_unique_containers(df, container_col):
    if not container_col or container_col not in df.columns: return []
    return sorted(df[container_col].dropna().astype(str).unique())

# --- CORRECTED Location Assignment Function ---
def automate_location_assignment(df, base_rack_id, rack_configs, status_text=None):
    # Standardize column names for processing
    required_cols = find_required_columns(df)
    
    if not all([required_cols['Part No'], required_cols['Container'], required_cols['Station No']]):
        st.error("❌ 'Part Number', 'Container Type', or 'Station No' column not found.")
        return None

    # Create a working copy with standardized names
    df_processed = df.copy()
    rename_dict = {v: k for k, v in required_cols.items() if v}
    df_processed.rename(columns=rename_dict, inplace=True)
    df_processed.sort_values(by=['Station No', 'Container'], inplace=True)

    final_parts_list = []
    
    for station_no, station_group in df_processed.groupby('Station No', sort=False):
        if status_text: status_text.text(f"Processing station: {station_no}...")

        rack_idx, level_idx = 0, 0
        sorted_racks = sorted(rack_configs.items())

        for container_type, parts_group in station_group.groupby('Container', sort=True):
            items_to_place = parts_group.to_dict('records')
            
            while items_to_place:
                slot_found = False
                search_rack_idx, search_level_idx = rack_idx, level_idx
                while search_rack_idx < len(sorted_racks):
                    rack_name, config = sorted_racks[search_rack_idx]
                    levels, capacity = config.get('levels', []), config.get('rack_bin_counts', {}).get(container_type, 0)

                    if capacity > 0 and search_level_idx < len(levels):
                        slot_found = True
                        rack_idx, level_idx = search_rack_idx, search_level_idx
                        break
                    
                    search_level_idx = 0
                    search_rack_idx += 1
                
                if not slot_found:
                    st.warning(f"⚠️ Ran out of rack space at Station {station_no} for '{container_type}'.")
                    break

                rack_name, config = sorted_racks[rack_idx]
                levels = config.get('levels', [])
                level_capacity = config.get('rack_bin_counts', {}).get(container_type, 0)

                parts_for_level = items_to_place[:level_capacity]
                items_to_place = items_to_place[level_capacity:]
                
                num_empty_slots = level_capacity - len(parts_for_level)
                level_items = parts_for_level + ([{'Part No': 'EMPTY'}] * num_empty_slots)
                
                # Create a template for empty items based on a real item's columns
                item_template = {col: '' for col in df_processed.columns}

                for cell_idx, item in enumerate(level_items, 1):
                    rack_num_val = ''.join(filter(str.isdigit, rack_name))
                    rack_num_1st = rack_num_val[0] if len(rack_num_val) > 1 else '0'
                    rack_num_2nd = rack_num_val[1] if len(rack_num_val) > 1 else rack_num_val[0]
                    
                    location_info = {
                        'Rack': base_rack_id, 'Rack No 1st': rack_num_1st, 'Rack No 2nd': rack_num_2nd,
                        'Level': levels[level_idx], 'Cell': str(cell_idx), 'Station No': station_no
                    }

                    if item['Part No'] == 'EMPTY':
                        full_item = item_template.copy()
                        full_item.update({'Part No': 'EMPTY', 'Container': container_type})
                    else:
                        full_item = item

                    full_item.update(location_info)
                    final_parts_list.append(full_item)

                level_idx += 1
                if level_idx >= len(levels):
                    level_idx = 0
                    rack_idx += 1
    
    if not final_parts_list: return pd.DataFrame()
    # The returned DataFrame now correctly represents every single slot.
    return pd.DataFrame(final_parts_list)

def create_location_key(row):
    return '_'.join([str(row.get(c, '')) for c in ['Station No', 'Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']])

def extract_location_values(row):
    return [str(row.get(c, '')) for c in ['Bus Model', 'Station No', 'Rack', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']]


# --- PDF Generation (Rack Labels - No changes needed, they now receive correct data) ---
def generate_rack_labels_v1(df, progress_bar=None, status_text=None):
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
        if status_text: status_text.text(f"Processing Rack Label {i+1}/{total_locations}")
        
        part1 = group.iloc[0]
        if str(part1['Part No']).upper() == 'EMPTY': continue

        rack_key = f"ST-{part1.get('Station No', 'NA')} / Rack {part1.get('Rack No 1st', '0')}{part1.get('Rack No 2nd', '0')}"
        label_summary[rack_key] = label_summary.get(rack_key, 0) + 1

        if label_count > 0 and label_count % 4 == 0: elements.append(PageBreak())
        
        part2 = group.iloc[1] if len(group) > 1 else part1
        
        part_table1 = Table([['Part No', format_part_no_v1(str(part1['Part No']))], ['Description', format_description_v1(str(part1['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.3*cm, 0.8*cm])
        part_table2 = Table([['Part No', format_part_no_v1(str(part2['Part No']))], ['Description', format_description_v1(str(part2['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.3*cm, 0.8*cm])
        
        location_values = extract_location_values(part1)
        location_data = [[Paragraph('Line Location', location_header_style)] + [Paragraph(str(val), location_value_style_v1) for val in location_values]]
        
        col_props = [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]
        location_widths = [4 * cm] + [w * (11 * cm) / sum(col_props) for w in col_props]
        location_table = Table(location_data, colWidths=location_widths, rowHeights=0.8*cm)
        
        part_style = TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (0, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, -1), 16)])
        part_table1.setStyle(part_style)
        part_table2.setStyle(part_style)
        
        loc_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
        loc_style_cmds = [('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]
        for j, color in enumerate(loc_colors):
            loc_style_cmds.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
        location_table.setStyle(TableStyle(loc_style_cmds))
        
        elements.extend([part_table1, Spacer(1, 0.3 * cm), part_table2, Spacer(1, 0.3 * cm), location_table, Spacer(1, 0.2 * cm)])
        label_count += 1
        
    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer, label_summary

def generate_rack_labels_v2(df, progress_bar=None, status_text=None):
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
        if status_text: status_text.text(f"Processing Rack Label {i+1}/{total_locations}")

        part1 = group.iloc[0]
        if str(part1['Part No']).upper() == 'EMPTY': continue
        
        rack_key = f"ST-{part1.get('Station No', 'NA')} / Rack {part1.get('Rack No 1st', '0')}{part1.get('Rack No 2nd', '0')}"
        label_summary[rack_key] = label_summary.get(rack_key, 0) + 1
            
        if label_count > 0 and label_count % 4 == 0: elements.append(PageBreak())

        part_table = Table([['Part No', format_part_no_v2(str(part1['Part No']))], ['Description', format_description(str(part1['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[1.9*cm, 2.1*cm])
        
        location_values = extract_location_values(part1)
        location_data = [[Paragraph('Line Location', location_header_style)] + [Paragraph(str(val), location_value_style_v2) for val in location_values]]

        col_widths = [1.7, 2.9, 1.3, 1.2, 1.3, 1.3, 1.3]
        location_widths = [4 * cm] + [w * (11 * cm) / sum(col_widths) for w in col_widths]
        location_table = Table(location_data, colWidths=location_widths, rowHeights=0.9*cm)
        
        part_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTER'), ('ALIGN', (1, 1), (1, -1), 'LEFT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 5), ('FONTNAME', (0, 0), (0, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, -1), 16)]))
        
        loc_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
        loc_style_cmds = [('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]
        for j, color in enumerate(loc_colors):
            loc_style_cmds.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
        location_table.setStyle(TableStyle(loc_style_cmds))
        
        elements.extend([part_table, Spacer(1, 0.3 * cm), location_table, Spacer(1, 0.2 * cm)])
        label_count += 1
        
    if elements: doc.build(elements)
    buffer.seek(0)
    return buffer, label_summary


# --- PDF Generation (Bin Labels Helpers) ---
def generate_qr_code_image(data_string):
    if not QR_AVAILABLE: return None
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4)
    qr.add_data(data_string)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")
    img_buffer = BytesIO()
    qr_img.save(img_buffer, format='PNG')
    img_buffer.seek(0)
    return Image(img_buffer, width=2.5*cm, height=2.5*cm)

def detect_bus_model_and_qty(row):
    result = {'7M': '', '9M': '', '12M': ''}
    qty_veh = str(row.get('Qty/Veh', ''))
    bus_model = str(row.get('Bus Model', '')).upper()
    
    if not qty_veh: return result

    detected_model = None
    if '7M' in bus_model or '7' == bus_model: detected_model = '7M'
    elif '9M' in bus_model or '9' == bus_model: detected_model = '9M'
    elif '12M' in bus_model or '12' == bus_model: detected_model = '12M'
    
    if detected_model:
        result[detected_model] = qty_veh
    return result

def extract_store_location_data_from_excel(row_data):
    col_lookup = {str(k).strip().upper(): k for k in row_data.keys()}

    def get_clean_value(possible_names, default=''):
        for name in possible_names:
            clean_name = name.strip().upper()
            if clean_name in col_lookup:
                original_col_name = col_lookup[clean_name]
                val = row_data.get(original_col_name)
                if pd.notna(val) and str(val).strip().lower() not in ['nan', 'none', 'null', '']:
                    return str(val).strip()
        return default

    store_location = get_clean_value(['Store Location', 'STORELOCATION', 'Store_Location'])
    zone = get_clean_value(['ABB ZONE', 'ABB_ZONE', 'ABBZONE'])
    location = get_clean_value(['ABB LOCATION', 'ABB_LOCATION', 'ABBLOCATION'])
    floor = get_clean_value(['ABB FLOOR', 'ABB_FLOOR', 'ABBFLOOR'])
    rack_no = get_clean_value(['ABB RACK NO', 'ABB_RACK_NO', 'ABBRACKNO'])
    level_in_rack = get_clean_value(['ABB LEVEL IN RACK', 'ABB_LEVEL_IN_RACK', 'ABBLEVELINRACK'])
    
    station_name = '' 
    return [station_name, store_location, zone, location, floor, rack_no, level_in_rack]

# --- PDF Generation (Bin Labels Main Function) ---
def generate_bin_labels(df, progress_bar=None, status_text=None):
    if not QR_AVAILABLE:
        st.error("❌ QR Code library not found. Please install `qrcode` and `Pillow`.")
        return None, {}

    STICKER_WIDTH, STICKER_HEIGHT = 10 * cm, 15 * cm
    CONTENT_BOX_WIDTH, CONTENT_BOX_HEIGHT = 10 * cm, 7.2 * cm
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=(STICKER_WIDTH, STICKER_HEIGHT),
                            topMargin=0.2*cm, bottomMargin=STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm,
                            leftMargin=0.1*cm, rightMargin=0.1*cm)

    df_filtered = df[df['Part No'].str.upper() != 'EMPTY'].copy()
    df_filtered.sort_values(by=['Station No', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell'], inplace=True)
    total_labels = len(df_filtered)
    label_summary = {}
    all_elements = []

    def draw_border(canvas, doc):
        canvas.saveState()
        x_offset = (STICKER_WIDTH - CONTENT_BOX_WIDTH) / 2
        y_offset = STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm
        canvas.setStrokeColorRGB(0, 0, 0)
        canvas.setLineWidth(1.8)
        canvas.rect(x_offset + doc.leftMargin, y_offset, CONTENT_BOX_WIDTH - 0.2*cm, CONTENT_BOX_HEIGHT)
        canvas.restoreState()

    for i, row in enumerate(df_filtered.to_dict('records')):
        if progress_bar: progress_bar.progress(int(((i+1) / total_labels) * 100))
        if status_text: status_text.text(f"Processing Bin Label {i+1}/{total_labels}")
        
        rack_key = f"ST-{row.get('Station No', 'NA')} / Rack {row.get('Rack No 1st', '0')}{row.get('Rack No 2nd', '0')}"
        label_summary[rack_key] = label_summary.get(rack_key, 0) + 1

        part_no = str(row.get('Part No', ''))
        desc = str(row.get('Description', ''))
        qty_bin = str(row.get('Qty/Bin', ''))

        qr_data = f"Part No: {part_no}\nDesc: {desc}\nLine Loc: {'_'.join(extract_location_values(row))}"
        qr_image = generate_qr_code_image(qr_data)
        
        content_width = CONTENT_BOX_WIDTH - 0.2*cm
        
        main_table = Table([
            ["Part No", Paragraph(f"{part_no}", bin_bold_style)],
            ["Description", Paragraph(desc[:47] + "..." if len(desc) > 50 else desc, bin_desc_style)],
            ["Qty/Bin", Paragraph(qty_bin, bin_qty_style)]
        ], colWidths=[content_width/3, content_width*2/3], rowHeights=[0.9*cm, 1.0*cm, 0.5*cm])
        main_table.setStyle(TableStyle([('GRID', (0,0),(-1,-1), 1.2, colors.black),('ALIGN', (0,0),(-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'MIDDLE'), ('FONTNAME', (0,0),(0,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0),(0,-1), 11)]))

        inner_table_width = content_width * 2 / 3
        col_props = [1.8, 2.4, 0.7, 0.7, 0.7, 0.7, 0.9]
        inner_col_widths = [w * inner_table_width / sum(col_props) for w in col_props]
        
        store_loc_values = extract_store_location_data_from_excel(row)
        store_loc_inner = Table([store_loc_values], colWidths=inner_col_widths, rowHeights=[0.5*cm])
        store_loc_inner.setStyle(TableStyle([('GRID', (0,0),(-1,-1), 1.2, colors.black), ('ALIGN', (0,0),(-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'MIDDLE'), ('FONTNAME', (0,0),(-1,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0),(-1,-1), 9)]))
        store_loc_table = Table([[Paragraph("Store Location", bin_desc_style), store_loc_inner]], colWidths=[content_width/3, inner_table_width], rowHeights=[0.5*cm])
        store_loc_table.setStyle(TableStyle([('GRID', (0,0),(-1,-1), 1.2, colors.black), ('ALIGN', (0,0),(-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'MIDDLE')]))
        
        line_loc_values = extract_location_values(row)
        line_loc_inner = Table([line_loc_values], colWidths=inner_col_widths, rowHeights=[0.5*cm])
        line_loc_inner.setStyle(TableStyle([('GRID', (0,0),(-1,-1), 1.2, colors.black), ('ALIGN', (0,0),(-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'MIDDLE'), ('FONTNAME', (0,0),(-1,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0),(-1,-1), 9)]))
        line_loc_table = Table([[Paragraph("Line Location", bin_desc_style), line_loc_inner]], colWidths=[content_width/3, inner_table_width], rowHeights=[0.5*cm])
        line_loc_table.setStyle(TableStyle([('GRID', (0,0),(-1,-1), 1.2, colors.black), ('ALIGN', (0,0),(-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'MIDDLE')]))

        mtm_quantities = detect_bus_model_and_qty(row)
        mtm_data = [
            ["7M", "9M", "12M"],
            [Paragraph(f"<b>{mtm_quantities['7M']}</b>", bin_qty_style) if mtm_quantities['7M'] else "",
             Paragraph(f"<b>{mtm_quantities['9M']}</b>", bin_qty_style) if mtm_quantities['9M'] else "",
             Paragraph(f"<b>{mtm_quantities['12M']}</b>", bin_qty_style) if mtm_quantities['12M'] else ""]
        ]
        mtm_table = Table(mtm_data, colWidths=[1.2*cm, 1.2*cm, 1.2*cm], rowHeights=[0.75*cm, 0.75*cm])
        mtm_table.setStyle(TableStyle([('GRID', (0,0),(-1,-1), 1.2, colors.black), ('ALIGN', (0,0),(-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'MIDDLE'), ('FONTNAME', (0,0),(-1,0), 'Helvetica-Bold'), ('FONTSIZE', (0,0),(-1,-1), 9)]))

        bottom_row = Table([[mtm_table, "", qr_image or ""]], colWidths=[3.6*cm, content_width - 3.6*cm - 2.5*cm, 2.5*cm], rowHeights=[2.5*cm])
        bottom_row.setStyle(TableStyle([('VALIGN', (0,0),(-1,-1), 'MIDDLE')]))

        all_elements.extend([main_table, store_loc_table, line_loc_table, Spacer(1, 0.2*cm), bottom_row])
        if i < total_labels - 1:
            all_elements.append(PageBreak())

    if all_elements: doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
    buffer.seek(0)
    return buffer, label_summary


# --- Main Application UI ---
def main():
    st.title("🏷️ AgiloSmartTag Studio")
    st.markdown("<p style='font-style:italic;'>Designed and Developed by Agilomatrix</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.sidebar.title("📄 Label Options")
    
    output_type = st.sidebar.selectbox("Choose Output Type:", ["Rack Labels", "Bin Labels"])

    rack_label_format = "Single Part"
    if output_type == "Rack Labels":
        rack_label_format = st.sidebar.selectbox("Choose Rack Label Format:", ["Single Part", "Multiple Parts"])

    base_rack_id = st.sidebar.text_input("Enter Storage Line Side Infrastructure", "R", help="E.g., R for Rack, TR for Tray.")
    st.sidebar.caption("EXAMPLE: **R** = RACK, **TR** = TRAY, **SH** = SHELVING")
    
    uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file, dtype=str) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, dtype=str)
            df.fillna('', inplace=True)
            st.success(f"✅ File loaded! Found {len(df)} rows.")
            
            required_cols_check = find_required_columns(df)
            
            if required_cols_check['Container']:
                unique_containers = get_unique_containers(df, required_cols_check['Container'])
                
                with st.expander("⚙️ Step 1: Configure Dimensions and Rack Setup (Applied to Each Station)", expanded=True):
                    
                    st.subheader("1. Container Dimensions")
                    bin_dims = {}
                    for container in unique_containers:
                        dim = st.text_input(f"Dimensions for {container}", key=f"bindim_{container}", placeholder="e.g., 300x200x150mm")
                        bin_dims[container] = dim
                    st.markdown("---")

                    st.subheader("2. Rack Dimensions & Bin/Level Capacity")
                    num_racks = st.number_input("Number of Racks (per station)", min_value=1, value=1, step=1)
                    
                    rack_configs = {}
                    rack_dims = {}
                    for i in range(num_racks):
                        rack_name = f"Rack {i+1:02d}"
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown(f"**Settings for {rack_name}**")
                            r_dim = st.text_input(f"Dimensions for {rack_name}", key=f"rackdim_{rack_name}", placeholder="e.g., 1200x1000x2000mm")
                            rack_dims[rack_name] = r_dim
                            levels = st.multiselect(f"Available Levels for {rack_name}",
                                options=['A','B','C','D','E','F','G','H'], default=['A','B','C','D','E'], key=f"levels_{rack_name}")
                        
                        with col2:
                            st.markdown(f"**Bin Capacity Per Level for {rack_name}**")
                            rack_bin_counts = {}
                            for container in unique_containers:
                                b_count = st.number_input(f"Capacity of '{container}' Bins", min_value=0, value=0, step=1, key=f"bcount_{rack_name}_{container}")
                                if b_count > 0: rack_bin_counts[container] = b_count
                        
                        rack_configs[rack_name] = {'dimensions': r_dim, 'levels': levels, 'rack_bin_counts': rack_bin_counts}
                        st.markdown("---")

                if st.button("🚀 Generate PDF Labels", type="primary"):
                    missing_bin_dims = [name for name, dim in bin_dims.items() if not dim]
                    missing_rack_dims = [name for name, dim in rack_dims.items() if not dim]
                    
                    error_messages = []
                    if missing_bin_dims: error_messages.append(f"container dimensions for: {', '.join(missing_bin_dims)}")
                    if missing_rack_dims: error_messages.append(f"rack dimensions for: {', '.join(missing_rack_dims)}")

                    if error_messages:
                        st.error(f"❌ Please provide all required information. Missing {'; '.join(error_messages)}.")
                    else:
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        try:
                            df_processed = automate_location_assignment(df, base_rack_id, rack_configs, status_text)
                            
                            if df_processed is not None and not df_processed.empty:
                                pdf_buffer, label_summary = None, {}
                                
                                if output_type == "Rack Labels":
                                    gen_func = generate_rack_labels_v2 if rack_label_format == "Single Part" else generate_rack_labels_v1
                                    pdf_buffer, label_summary = gen_func(df_processed, progress_bar, status_text)
                                elif output_type == "Bin Labels":
                                    pdf_buffer, label_summary = generate_bin_labels(df_processed, progress_bar, status_text)

                                if pdf_buffer and sum(label_summary.values()) > 0:
                                    total_labels = sum(label_summary.values())
                                    status_text.text(f"✅ PDF with {total_labels} labels generated successfully!")
                                    file_name_suffix = "rack_labels.pdf" if output_type == "Rack Labels" else "bin_labels.pdf"
                                    file_name = f"{os.path.splitext(uploaded_file.name)[0]}_{file_name_suffix}"
                                    st.download_button(label="📥 Download PDF", data=pdf_buffer.getvalue(), file_name=file_name, mime="application/pdf")

                                    st.markdown("---")
                                    st.subheader("📊 Generation Summary")
                                    summary_df = pd.DataFrame(list(label_summary.items()), columns=['Location', 'Number of Labels']).sort_values(by='Location').reset_index(drop=True)
                                    st.table(summary_df)
                                else:
                                    st.warning("⚠️ No labels were generated. This could be due to no parts in the input file or rack capacity being zero.")
                            else:
                                st.error("❌ No data was processed. Please check the input file and configurations.")
                        except Exception as e:
                            st.error(f"❌ An unexpected error occurred: {e}")
                            st.exception(e)
                        finally:
                            progress_bar.empty()
                            status_text.empty()
            else:
                st.error("❌ A column containing 'Container' is required and could not be found in the uploaded file.")
        except Exception as e:
            st.error(f"❌ Error reading file: {e}")
    else:
        st.info("👆 Upload a file to begin.")

if __name__ == "__main__":
    main()
