import streamlit as st
import pandas as pd
import os
import re
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# Configure Streamlit page
st.set_page_config(
    page_title="Part Label Generator",
    page_icon="🏷️",
    layout="wide"
)

# Style definitions (Your provided styles)
bold_style_v1 = ParagraphStyle(
    name='Bold_v1',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_LEFT,
    leading=20,
    spaceBefore=2,
    spaceAfter=2
)

bold_style_v2 = ParagraphStyle(
    name='Bold_v2',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_LEFT,
    leading=12,
    spaceBefore=0,
    spaceAfter=15,
)

desc_style = ParagraphStyle(
    name='Description',
    fontName='Helvetica',
    fontSize=20,
    alignment=TA_LEFT,
    leading=16,
    spaceBefore=2,
    spaceAfter=2
)

# Formatting functions (Your provided functions)
def format_part_no_v1(part_no):
    if not part_no or not isinstance(part_no, str):
        part_no = str(part_no)
    if len(part_no) > 5:
        split_point = len(part_no) - 5
        part1 = part_no[:split_point]
        part2 = part_no[-5:]
        return Paragraph(f"<b><font size=17>{part1}</font><font size=22>{part2}</font></b>", bold_style_v1)
    else:
        return Paragraph(f"<b><font size=17>{part_no}</font></b>", bold_style_v1)

def format_part_no_v2(part_no):
    if not part_no or not isinstance(part_no, str):
        part_no = str(part_no)
    if len(part_no) > 5:
        split_point = len(part_no) - 5
        part1 = part_no[:split_point]
        part2 = part_no[-5:]
        return Paragraph(f"<b><font size=34>{part1}</font><font size=40>{part2}</font></b><br/><br/>", bold_style_v2)
    else:
        return Paragraph(f"<b><font size=34>{part_no}</font></b><br/><br/>", bold_style_v2)

def format_description_v1(desc):
    if not desc or not isinstance(desc, str):
        desc = str(desc)
    desc_length = len(desc)
    if desc_length <= 30: font_size = 15
    elif desc_length <= 50: font_size = 13
    elif desc_length <= 70: font_size = 11
    elif desc_length <= 90: font_size = 10
    else:
        font_size = 9
        desc = desc[:100] + "..." if len(desc) > 100 else desc
    desc_style_v1 = ParagraphStyle(
        name='Description_v1', fontName='Helvetica', fontSize=font_size, alignment=TA_LEFT, leading=font_size + 2, spaceBefore=1, spaceAfter=1
    )
    return Paragraph(desc, desc_style_v1)

def format_description(desc):
    if not desc or not isinstance(desc, str):
        desc = str(desc)
    return Paragraph(desc, desc_style)

# Core Logic and Helper Functions (Your provided functions)
def extract_location_values(row, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col):
    location_values = [''] * 7
    location_values[0] = str(row.get(bus_model_col, '')) if bus_model_col and bus_model_col in row else ''
    location_values[1] = str(row.get(station_no_col, '')) if station_no_col and station_no_col in row else ''
    location_values[2] = str(row.get(rack_col, '')) if rack_col and rack_col in row else ''
    if rack_no_1st_col and rack_no_1st_col in row:
        location_values[3] = str(row.get(rack_no_1st_col, ''))
    elif rack_no_col and rack_no_col in row:
        rack_no_value = str(row.get(rack_no_col, ''))
        location_values[3] = rack_no_value[0] if rack_no_value and len(rack_no_value) >= 1 else ''
    if rack_no_2nd_col and rack_no_2nd_col in row:
        location_values[4] = str(row.get(rack_no_2nd_col, ''))
    elif rack_no_col and rack_no_col in row:
        rack_no_value = str(row.get(rack_no_col, ''))
        location_values[4] = rack_no_value[1] if rack_no_value and len(rack_no_value) >= 2 else ''
    location_values[5] = str(row.get(level_col, '')) if level_col and level_col in row else ''
    location_values[6] = str(row.get(cell_col, '')) if cell_col and cell_col in row else ''
    return location_values

def find_location_columns(df):
    cols = [str(col).upper() for col in df.columns]
    original_cols = df.columns.tolist()
    
    def find_col(patterns):
        for pattern_list in patterns:
            for i, col_name in enumerate(cols):
                if all(p in col_name for p in pattern_list):
                    return original_cols[i]
        return None

    bus_model_col = find_col([['BUS', 'MODEL'], ['BUS']])
    station_no_col = find_col([['STATION', 'NO'], ['STATION', 'NUM'], ['STATION']])
    rack_col = next((original_cols[i] for i, col in enumerate(cols) if 'RACK' in col and 'NO' not in col), None)
    rack_no_1st_col = find_col([['RACK', 'NO', '1ST'], ['RACK', '1', 'DIGIT']])
    rack_no_2nd_col = find_col([['RACK', 'NO', '2ND'], ['RACK', '2', 'DIGIT']])
    rack_no_col = find_col([['RACK', 'NO'], ['RACK', 'NUM']])
    level_col = find_col([['LEVEL']])
    cell_col = find_col([['CELL']])
    
    return bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col

def create_location_key(row, cols_map):
    values = [str(row.get(cols_map[key], '')) for key in ['bus', 'station', 'rack', 'level', 'cell']]
    
    # Handle rack_no separately
    rack_no_val = ''
    if cols_map.get('rack_no_1st'):
        rack_no_val += str(row.get(cols_map['rack_no_1st'], ''))
    if cols_map.get('rack_no_2nd'):
        rack_no_val += str(row.get(cols_map['rack_no_2nd'], ''))
    elif cols_map.get('rack_no'): # Fallback to single column
        rack_no_full = str(row.get(cols_map['rack_no'], ''))
        rack_no_val = rack_no_full.ljust(2, ' ') # Pad to 2 chars to avoid index errors
    
    values.insert(3, rack_no_val)
    return '_'.join(values)

# PDF Generation (UPDATED WITH SPACING FIX)
def generate_labels_from_excel_v1(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    part_no_height, desc_loc_height = 1.3 * cm, 0.8 * cm

    # Find columns
    part_no_col = next((c for c in df.columns if 'PART' in str(c).upper() and ('NO' in str(c).upper() or 'NUM' in str(c).upper())), df.columns[0])
    desc_col = next((c for c in df.columns if 'DESC' in str(c).upper()), df.columns[1])
    bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col = find_location_columns(df)
    
    cols_map = {
        'bus': bus_model_col, 'station': station_no_col, 'rack': rack_col,
        'rack_no': rack_no_col, 'rack_no_1st': rack_no_1st_col, 'rack_no_2nd': rack_no_2nd_col,
        'level': level_col, 'cell': cell_col
    }

    df['location_key'] = df.apply(create_location_key, axis=1, cols_map=cols_map)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)

    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.5*cm, bottomMargin=1*cm, leftMargin=2*cm, rightMargin=2*cm)
    elements = []
    MAX_LABELS_PER_PAGE = 4
    label_count = 0

    for i, (location_key, group) in enumerate(df_grouped):
        try:
            if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
            if status_text: status_text.text(f"Processing location {i+1}/{total_locations}")

            part1 = group.iloc[0]
            part2 = group.iloc[1] if len(group) > 1 else part1
                
            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())
            label_count += 1

            location_values = extract_location_values(part1, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col)
            
            part_table = Table([['Part No', format_part_no_v1(str(part1[part_no_col]))], ['Description', format_description_v1(str(part1[desc_col]))]], colWidths=[4*cm, 11*cm], rowHeights=[part_no_height, desc_loc_height])
            part_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('ALIGN', (0,0), (0,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (0,-1), 16)]))
            
            part_table2 = Table([['Part No', format_part_no_v1(str(part2[part_no_col]))], ['Description', format_description_v1(str(part2[desc_col]))]], colWidths=[4*cm, 11*cm], rowHeights=[part_no_height, desc_loc_height])
            part_table2.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('ALIGN', (0,0), (0,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (0,-1), 16)]))

            location_data = [['Line Location'] + location_values]
            col_proportions = [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]
            location_widths = [4 * cm] + [w * (11 * cm) / sum(col_proportions) for w in col_proportions]
            location_table = Table(location_data, colWidths=location_widths, rowHeights=desc_loc_height)
            location_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
            location_style = [('GRID', (0,0), (-1,-1), 1, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (0,0), 16), ('FONTSIZE', (1,0), (-1,-1), 14)]
            for j, color in enumerate(location_colors): location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
            location_table.setStyle(TableStyle(location_style))
            
            elements.append(part_table)
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(part_table2)
            # --- FIX: Added Spacer for vertical gap ---
            elements.append(Spacer(1, 0.3 * cm)) 
            elements.append(location_table)
            elements.append(Spacer(1, 0.2 * cm))

        except Exception as e:
            if status_text: status_text.text(f"Error at location {location_key}: {e}")
            continue

    if elements:
        doc.build(elements)
        buffer.seek(0)
        return buffer
    return None

def generate_labels_from_excel_v2(df, progress_bar=None, status_text=None):
    buffer = io.BytesIO()
    part_no_height, desc_height, loc_height = 1.9 * cm, 2.1 * cm, 0.9 * cm

    # Find columns
    part_no_col = next((c for c in df.columns if 'PART' in str(c).upper() and ('NO' in str(c).upper() or 'NUM' in str(c).upper())), df.columns[0])
    desc_col = next((c for c in df.columns if 'DESC' in str(c).upper()), df.columns[1])
    bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col = find_location_columns(df)
    
    cols_map = {
        'bus': bus_model_col, 'station': station_no_col, 'rack': rack_col,
        'rack_no': rack_no_col, 'rack_no_1st': rack_no_1st_col, 'rack_no_2nd': rack_no_2nd_col,
        'level': level_col, 'cell': cell_col
    }
    
    df['location_key'] = df.apply(create_location_key, axis=1, cols_map=cols_map)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)
    
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.5*cm, bottomMargin=1*cm, leftMargin=2*cm, rightMargin=2*cm)
    elements = []
    MAX_LABELS_PER_PAGE = 4
    label_count = 0

    for i, (location_key, group) in enumerate(df_grouped):
        try:
            if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
            if status_text: status_text.text(f"Processing location {i+1}/{total_locations}")

            part1 = group.iloc[0]
            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())
            label_count += 1

            location_values = extract_location_values(part1, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col)
            
            part_table = Table([['Part No', format_part_no_v2(str(part1[part_no_col]))], ['Description', format_description(str(part1[desc_col]))]], colWidths=[4*cm, 11*cm], rowHeights=[part_no_height, desc_height])
            part_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('ALIGN', (0,0), (0,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (0,-1), 16), ('LEFTPADDING', (1,1), (1,1), 10)]))

            location_data = [['Line Location'] + location_values]
            col_widths = [1.7, 2.9, 1.3, 1.2, 1.3, 1.3, 1.3]
            location_widths = [4 * cm] + [w * (11 * cm) / sum(col_widths) for w in col_widths]
            location_table = Table(location_data, colWidths=location_widths, rowHeights=loc_height)
            location_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
            location_style = [('GRID', (0,0), (-1,-1), 1, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (0,0), 16), ('FONTSIZE', (1,0), (-1,-1), 16)]
            for j, color in enumerate(location_colors): location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
            location_table.setStyle(TableStyle(location_style))
            
            elements.append(part_table)
            # --- This Spacer was already here and correct ---
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(location_table)
            elements.append(Spacer(1, 0.2 * cm))

        except Exception as e:
            if status_text: status_text.text(f"Error at location {location_key}: {e}")
            continue

    if elements:
        doc.build(elements)
        buffer.seek(0)
        return buffer
    return None

# Main App UI (Your provided UI)
def main():
    st.title("🏷️ Rack Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )
    st.markdown("---")
    st.sidebar.title("Label Generator Options")
    label_type = st.sidebar.selectbox("Choose Rack Type:", ["Single Part", "Multiple Parts"])
    uploaded_file = st.file_uploader(
        "Choose an Excel or CSV file", type=['xlsx', 'xls', 'csv'], help="Upload your Excel or CSV file containing part information"
    )

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith('.csv') else pd.read_excel(uploaded_file)
            st.success(f"✅ File loaded successfully! Found {len(df)} rows.")
            
            with st.expander("📊 File Information", expanded=False):
                st.write("**Columns found:**", df.columns.tolist())
                st.dataframe(df.head(3))

            if st.button("🚀 Generate PDF Labels", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                try:
                    if label_type == "Single Part":
                        pdf_buffer = generate_labels_from_excel_v2(df, progress_bar, status_text)
                        filename = f"{os.path.splitext(uploaded_file.name)[0]}_single_part_labels.pdf"
                    else:
                        pdf_buffer = generate_labels_from_excel_v1(df, progress_bar, status_text)
                        filename = f"{os.path.splitext(uploaded_file.name)[0]}_multi_part_labels.pdf"
                    
                    if pdf_buffer:
                        status_text.text("✅ PDF generated successfully!")
                        st.download_button(
                            label="📥 Download PDF Labels", data=pdf_buffer.getvalue(), file_name=filename, mime="application/pdf"
                        )
                        # Display stats
                        with st.expander("📈 Generation Statistics", expanded=True):
                            cols_map = dict(zip(['bus', 'station', 'rack', 'rack_no', 'rack_no_1st', 'rack_no_2nd', 'level', 'cell'], find_location_columns(df)))
                            unique_locations = df.apply(create_location_key, axis=1, cols_map=cols_map).nunique()
                            st.metric("Total Parts Processed", len(df))
                            st.metric("Unique Locations Found", unique_locations)
                            st.metric("Labels Generated", unique_locations)
                    else:
                        st.error("❌ Failed to generate PDF. Check file format and columns.")
                except Exception as e:
                    st.error(f"❌ Error during PDF generation: {str(e)}")
                finally:
                    progress_bar.empty()
                    status_text.empty()
        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")
    else:
        st.info("👆 Please upload a file to begin.")
        with st.expander("📋 Instructions", expanded=True):
            st.markdown("""
            ### How to use this tool:
            1.  **Select Label Format**: Choose between **Single Part** or **Multiple Parts**.
            2.  **Upload your file**: Choose an Excel or CSV file.
            3.  **Generate & Download**: Click the generate button to create and download your PDF.
            
            #### Required Columns in Your File:
            The tool will automatically try to find columns with names like:
            -   `PART NO`, `Part Number`, `DESC`, `DESCRIPTION`
            -   `Bus Model`, `Station No`, `Rack`, `Level`, `Cell`
            -   `Rack No` or separate `Rack No 1st` and `Rack No 2nd`
            """)

if __name__ == "__main__":
    main()
