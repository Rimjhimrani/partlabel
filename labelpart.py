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

# --- Style Definitions (No Changes Here) ---
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


# --- Formatting Functions (No Changes Here) ---
def format_part_no_v1(part_no):
    """Format part number with first 7 characters in 17pt font, rest in 22pt font."""
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
    """Format part number with different font sizes to prevent overlapping."""
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
    """Format description text with dynamic font sizing based on length for v1."""
    if not desc or not isinstance(desc, str):
        desc = str(desc)
    
    desc_length = len(desc)
    
    if desc_length <= 30:
        font_size = 15
    elif desc_length <= 50:
        font_size = 13
    elif desc_length <= 70:
        font_size = 11
    elif desc_length <= 90:
        font_size = 10
    else:
        font_size = 9
        desc = desc[:100] + "..." if len(desc) > 100 else desc
    
    desc_style_v1 = ParagraphStyle(
        name='Description_v1',
        fontName='Helvetica',
        fontSize=font_size,
        alignment=TA_LEFT,
        leading=font_size + 2,
        spaceBefore=1,
        spaceAfter=1
    )
    
    return Paragraph(desc, desc_style_v1)

def format_description(desc):
    """Format description text with proper wrapping."""
    if not desc or not isinstance(desc, str):
        desc = str(desc)
    return Paragraph(desc, desc_style)

# --- NEW AND UPDATED Core Logic Functions ---

def find_required_columns(df):
    """Find essential columns in the DataFrame for processing."""
    cols = {col.upper(): col for col in df.columns}
    
    part_no_key = next((k for k in cols if 'PART' in k and ('NO' in k or 'NUM' in k or '#' in k)), 
                       next((k for k in cols if k in ['PARTNO', 'PART']), None))
    desc_key = next((k for k in cols if 'DESC' in k), None)
    bus_model_key = next((k for k in cols if 'BUS' in k and 'MODEL' in k), 
                         next((k for k in cols if 'MODEL' in k), None))
    station_no_key = next((k for k in cols if 'STATION' in k and ('NO' in k or 'NUM' in k)), 
                          next((k for k in cols if 'STATION' in k), None))
    container_type_key = next((k for k in cols if 'CONTAINER' in k), None)

    return (cols.get(part_no_key), cols.get(desc_key), cols.get(bus_model_key), 
            cols.get(station_no_key), cols.get(container_type_key))

def get_unique_bins(df, container_col):
    """Finds and sorts unique container types containing 'BIN'."""
    if not container_col or container_col not in df.columns:
        return []
    
    # Extract unique values that contain 'bin' (case-insensitive), then sort them
    unique_containers = df[container_col].dropna().astype(str)
    bins = sorted([c for c in unique_containers.unique() if 'BIN' in c.upper()])
    return bins

def process_and_assign_locations(df, rack_input, level_selections_by_bin, status_text=None):
    """Processes the DataFrame to assign automated location values based on user input."""
    
    part_no_col, desc_col, model_col, station_col, container_col = find_required_columns(df)
    
    if not part_no_col or not container_col:
        st.error("❌ Critical columns 'Part Number' or 'Container Type' could not be found.")
        return None
        
    if status_text:
        status_text.text(f"Automating with columns: Part No='{part_no_col}', Container='{container_col}'")

    # --- NEW: Dynamic Rack Numbering based on sorted bin types ---
    sorted_unique_bins = get_unique_bins(df, container_col)
    bin_to_rack_num = {bin_name: index + 1 for index, bin_name in enumerate(sorted_unique_bins)}

    processed_rows = []
    level_counters = {} # To cycle through user-selected levels for each container type

    for _, row in df.iterrows():
        container_type = str(row.get(container_col, '')).strip()
        
        # --- Automated Location Logic ---
        rack_no_1st = ''
        rack_no_2nd = ''
        
        # Get sequential rack number based on the bin type
        rack_num = bin_to_rack_num.get(container_type)
        if rack_num is not None:
            rack_num_str = f"{rack_num:02d}"  # Format as two digits (e.g., 1 -> "01", 12 -> "12")
            rack_no_1st = rack_num_str[0]
            rack_no_2nd = rack_num_str[1]

        # Get the specific levels chosen by the user for THIS container type
        level_options = level_selections_by_bin.get(container_type, [])
        
        if level_options:
            level_idx = level_counters.get(container_type, 0)
            assigned_level = level_options[level_idx]
            level_counters[container_type] = (level_idx + 1) % len(level_options)
        else:
            assigned_level = ''

        new_row = {
            'Part No': row.get(part_no_col, ''),
            'Description': row.get(desc_col, ''),
            'Bus Model': row.get(model_col, ''),
            'Station No': row.get(station_col, ''),
            'Rack': rack_input,
            'Rack No 1st': rack_no_1st,
            'Rack No 2nd': rack_no_2nd,
            'Level': assigned_level,
            'Cell': '' # Cell is kept empty as requested
        }
        processed_rows.append(new_row)
        
    return pd.DataFrame(processed_rows)


def create_location_key(row):
    """Create a unique key for grouping by the newly generated location."""
    return '_'.join([
        str(row.get('Bus Model', '')),
        str(row.get('Station No', '')),
        str(row.get('Rack', '')),
        str(row.get('Rack No 1st', '')),
        str(row.get('Rack No 2nd', '')),
        str(row.get('Level', '')),
        str(row.get('Cell', ''))
    ])

def extract_location_values(row):
    """Extract location values from the processed row's columns."""
    return [
        str(row.get('Bus Model', '')),
        str(row.get('Station No', '')),
        str(row.get('Rack', '')),
        str(row.get('Rack No 1st', '')),
        str(row.get('Rack No 2nd', '')),
        str(row.get('Level', '')),
        str(row.get('Cell', ''))
    ]

# --- PDF Generation Functions (Updated to use new logic) ---

def generate_labels_from_excel_v1(df, progress_bar=None, status_text=None):
    """Generate labels using version 1 formatting (Multi-Part)."""
    
    buffer = io.BytesIO()
    part_no_height, desc_loc_height = 1.3 * cm, 0.8 * cm

    df['location_key'] = df.apply(create_location_key, axis=1)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)

    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    MAX_LABELS_PER_PAGE = 4
    label_count = 0

    for i, (location_key, group) in enumerate(df_grouped):
        try:
            if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
            if status_text: status_text.text(f"Processing location {i+1}/{total_locations}: {location_key.replace('_', ' ')}")

            parts = group.head(2)
            if len(parts) == 0: continue
            
            part1 = parts.iloc[0]
            part2 = parts.iloc[1] if len(parts) > 1 else part1
                
            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())
            label_count += 1

            location_values = extract_location_values(part1)

            part_table = Table([['Part No', format_part_no_v1(str(part1['Part No']))], ['Description', format_description_v1(str(part1['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[part_no_height, desc_loc_height])
            part_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTRE'), ('ALIGN', (1, 0), (1, -1), 'LEFT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 5), ('RIGHTPADDING', (0, 0), (-1, -1), 5), ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3), ('FONTNAME', (0, 0), (0, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, -1), 16)]))

            part_table2 = Table([['Part No', format_part_no_v1(str(part2['Part No']))], ['Description', format_description_v1(str(part2['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[part_no_height, desc_loc_height])
            part_table2.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTRE'), ('ALIGN', (1, 0), (1, -1), 'LEFT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 5), ('RIGHTPADDING', (0, 0), (-1, -1), 5), ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3), ('FONTNAME', (0, 0), (0, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, -1), 16)]))

            location_data = [['Part Location'] + location_values]
            col_proportions = [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]
            location_widths = [4 * cm] + [w * (11 * cm) / sum(col_proportions) for w in col_proportions]
            
            location_table = Table(location_data, colWidths=location_widths, rowHeights=desc_loc_height)
            location_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
            location_style = [('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, 0), 16), ('FONTSIZE', (1, 0), (-1, -1), 14)]
            for j, color in enumerate(location_colors):
                location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
            location_table.setStyle(TableStyle(location_style))
            
            elements.extend([part_table, Spacer(1, 0.3 * cm), part_table2, location_table, Spacer(1, 0.2 * cm)])
        except Exception as e:
            if status_text: status_text.text(f"Error processing location {location_key}: {e}")
            continue

    if progress_bar: progress_bar.progress(100)
    if elements:
        if status_text: status_text.text("Building PDF document...")
        doc.build(elements)
        buffer.seek(0)
        return buffer
    return None

def generate_labels_from_excel_v2(df, progress_bar=None, status_text=None):
    """Generate labels using version 2 formatting (Single Part)."""
    
    buffer = io.BytesIO()
    part_no_height, desc_height, loc_height = 1.9 * cm, 2.1 * cm, 0.9 * cm

    df['location_key'] = df.apply(create_location_key, axis=1)
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)

    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    MAX_LABELS_PER_PAGE = 4
    label_count = 0

    for i, (location_key, group) in enumerate(df_grouped):
        try:
            if progress_bar: progress_bar.progress(int((i / total_locations) * 100))
            if status_text: status_text.text(f"Processing location {i+1}/{total_locations}: {location_key.replace('_', ' ')}")

            part1 = group.iloc[0]
            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())
            label_count += 1

            location_values = extract_location_values(part1)

            part_table = Table([['Part No', format_part_no_v2(str(part1['Part No']))], ['Description', format_description(str(part1['Description']))]], colWidths=[4*cm, 11*cm], rowHeights=[part_no_height, desc_height])
            part_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (0, -1), 'CENTER'), ('ALIGN', (1, 0), (1, 0), 'CENTER'), ('ALIGN', (1, 1), (1, -1), 'LEFT'), ('VALIGN', (0, 0), (0, 0), 'MIDDLE'), ('VALIGN', (1, 0), (1, 0), 'TOP'), ('VALIGN', (0, 1), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 5), ('RIGHTPADDING', (0, 0), (-1, -1), 5), ('TOPPADDING', (1, 0), (1, 0), 10), ('BOTTOMPADDING', (1, 0), (1, 0), 5), ('FONTNAME', (0, 0), (0, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, -1), 16)]))
            
            location_data = [['Part Location'] + location_values]
            col_widths = [1.7, 2.9, 1.3, 1.2, 1.3, 1.3, 1.3]
            location_widths = [4 * cm] + [w * (11 * cm) / sum(col_widths) for w in col_widths]

            location_table = Table(location_data, colWidths=location_widths, rowHeights=loc_height)
            location_colors = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
            location_style = [('GRID', (0, 0), (-1, -1), 1, colors.black), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (0, 0), 16), ('FONTSIZE', (1, 0), (-1, -1), 16)]
            for j, color in enumerate(location_colors):
                location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))
            location_table.setStyle(TableStyle(location_style))
            
            elements.extend([part_table, Spacer(1, 0.3 * cm), location_table, Spacer(1, 0.2 * cm)])
        except Exception as e:
            if status_text: status_text.text(f"Error processing location {location_key}: {e}")
            continue

    if progress_bar: progress_bar.progress(100)
    if elements:
        if status_text: status_text.text("Building PDF document...")
        doc.build(elements)
        buffer.seek(0)
        return buffer
    return None


def main():
    st.title("🏷️ Rack Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )

    st.markdown("---")

    st.sidebar.title("Label Generator Options")
    label_type = st.sidebar.selectbox(
        "Choose Label Format:",
        ["Single Part", "Multiple Parts"],
        help="Single Part: One part per location label. Multiple Parts: Up to two parts per location label."
    )

    st.sidebar.title("Static Location Settings")
    rack_input = st.sidebar.text_input(
        "Enter Rack Value", 
        "TR",
        help="Enter the value for the 'Rack' field (e.g., TR, R, S)."
    )
    
    uploaded_file = st.file_uploader(
        "Choose an Excel or CSV file",
        type=['xlsx', 'xls', 'csv'],
        help="Upload your file with Part No, Description, Model, Station No, and Container Type."
    )

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith('.csv') else pd.read_excel(uploaded_file)
            st.success(f"✅ File loaded successfully! Found {len(df)} rows.")
            
            with st.expander("📊 File Preview", expanded=False):
                st.dataframe(df.head(3))

            # --- NEW: Dynamic UI for Level Selection ---
            level_selections = {}
            _, _, _, _, container_col = find_required_columns(df)
            
            if container_col:
                unique_bins = get_unique_bins(df, container_col)
                if unique_bins:
                    st.sidebar.title("Level Automation Settings")
                    st.sidebar.info("Select the levels to assign for each detected bin type.")
                    for bin_type in unique_bins:
                        level_selections[bin_type] = st.sidebar.multiselect(
                            f"Levels for {bin_type}",
                            options=['A', 'B', 'C', 'D', 'E'],
                            default=['A', 'B', 'C', 'D', 'E'],
                            key=bin_type 
                        )
                else:
                    st.warning("No 'Bin' types found in the 'Container Type' column to automate levels.")
            else:
                st.error("Could not find a 'Container Type' column in the file. Level automation is disabled.")


            if st.button("🚀 Generate PDF Labels", type="primary"):
                # Validation check
                is_ready = True
                if not rack_input:
                    st.warning("Please enter a value for the Rack.")
                    is_ready = False
                if container_col and not level_selections:
                    st.warning("No bin types were detected or configured for level automation.")
                    # This is not a blocking error, can proceed with empty levels.

                if is_ready:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    try:
                        status_text.text("Automating part locations...")
                        df_processed = process_and_assign_locations(df, rack_input, level_selections, status_text)
                        
                        if df_processed is not None:
                            if label_type == "Single Part":
                                pdf_buffer = generate_labels_from_excel_v2(df_processed, progress_bar, status_text)
                                filename = "singlepart_labels.pdf"
                            else:
                                pdf_buffer = generate_labels_from_excel_v1(df_processed, progress_bar, status_text)
                                filename = "multipart_labels.pdf"
                            
                            if pdf_buffer:
                                status_text.text("✅ PDF generated successfully!")
                                st.download_button(label="📥 Download PDF Labels", data=pdf_buffer.getvalue(), file_name=filename, mime="application/pdf")
                                with st.expander("📈 Generation Statistics", expanded=True):
                                    unique_locations = df_processed['location_key'].nunique()
                                    st.metric("Total Parts Processed", len(df_processed))
                                    st.metric("Unique Locations Created", unique_locations)
                                    st.metric("Labels Generated", unique_locations)
                            else:
                                st.error("❌ Failed to generate PDF. Check file data and columns.")
                                
                    except Exception as e:
                        st.error(f"❌ An error occurred: {str(e)}")
                    finally:
                        progress_bar.empty()
                        status_text.empty()

        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")

    else:
        st.info("👆 Upload a file to configure the automation settings and begin.")
        
        with st.expander("📋 Instructions", expanded=True):
            st.markdown("""
            ### How to use this tool:
            1. **Set Static Settings**: In the sidebar, enter the constant **Rack** value (e.g., TR).
            2. **Upload your file**: Choose an Excel or CSV file.
            3. **Configure Dynamic Levels**: The tool will detect bin types (Bin A, Bin B...) and show new options in the sidebar. Select the levels you want to assign to each bin type.
            4. **Select Label Format**: Choose between **Single Part** or **Multiple Parts** per label.
            5. **Generate & Download**: Click the generate button to create and download your PDF.
            
            ### Required Columns in Your File:
            - **Part Number** (e.g., "PART NO")
            - **Container Type** (e.g., "CONTAINER TYPE"). **Crucially, this must contain values like "Bin A", "Bin B", etc., for the automation to work.**
            - *Optional but recommended:* Description, Model, Station No.
            """)

if __name__ == "__main__":
    main()
