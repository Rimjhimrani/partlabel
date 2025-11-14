import streamlit as st
import pandas as pd
import os
import math
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

# --- Core Logic Functions (UPDATED) ---

def find_required_columns(df):
    """Find essential columns in the DataFrame for processing."""
    cols = {col.upper().strip(): col for col in df.columns}
    
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
    
    unique_containers = df[container_col].dropna().astype(str)
    bins = sorted([c for c in unique_containers.unique() if 'BIN' in c.upper()])
    return bins

def automate_location_assignment(df, rack_input, level_selections, bin_capacities, standard_bin_dimensions, level_dimensions_by_level, status_text=None):
    """
    Processes the DataFrame to assign automated location values based on dimensions,
    bin type, capacity, and level selections.
    """
    part_no_col, _, _, _, container_col = find_required_columns(df)

    if not part_no_col or not container_col:
        st.error("❌ Critical columns for 'Part Number' or 'Container Type' could not be found.")
        return None

    if status_text:
        status_text.text("Calculating bin placement based on dimensions...")

    # --- Pre-calculate how many bins fit on each unique level ---
    bins_per_level_config = {}
    bin_dim = standard_bin_dimensions
    
    for level, level_dim in level_dimensions_by_level.items():
        if not bin_dim or not level_dim or bin_dim.get('L', 0) == 0 or bin_dim.get('W', 0) == 0:
            bins_that_fit = 1
        else:
            fit_l = math.floor(level_dim['L'] / bin_dim['L']) if bin_dim['L'] > 0 else 0
            fit_w = math.floor(level_dim['W'] / bin_dim['W']) if bin_dim['W'] > 0 else 0
            bins_that_fit = max(1, fit_l * fit_w)
        bins_per_level_config[level] = bins_that_fit
        
    # --- Initialize state trackers and output lists ---
    df_sorted = df.sort_values(by=container_col).reset_index(drop=True)
    
    rack_num_for_bin_type = {bin_type: 1 for bin_type in get_unique_bins(df, container_col)}
    part_counters = {}  # Key: (bin_type, rack_num, level, cell_num), Value: count
    
    rack_list, r1_list, r2_list, lvl_list, cell_list = [], [], [], [], []

    # --- Main Loop: Iterate through each part and find a spot ---
    for _, row in df_sorted.iterrows():
        bin_type = row[container_col]
        
        assigned_levels = level_selections.get(bin_type, [])
        part_capacity_per_bin = bin_capacities.get(bin_type, 1)

        if not assigned_levels:
            rack_list.append(rack_input); r1_list.append(''); r2_list.append(''); lvl_list.append(''); cell_list.append('')
            continue

        found_spot = False
        while not found_spot:
            current_rack = rack_num_for_bin_type.get(bin_type, 1)
            
            for level in assigned_levels:
                max_bins_on_level = bins_per_level_config.get(level, 1)
                
                for cell in range(1, max_bins_on_level + 1):
                    key = (bin_type, current_rack, level, cell)
                    parts_in_this_cell = part_counters.get(key, 0)
                    
                    if parts_in_this_cell < part_capacity_per_bin:
                        part_counters[key] = parts_in_this_cell + 1
                        rack_num_str = f"{current_rack:02d}"
                        rack_list.append(rack_input)
                        r1_list.append(rack_num_str[0])
                        r2_list.append(rack_num_str[1])
                        lvl_list.append(level)
                        cell_list.append(f"{cell:02d}")
                        found_spot = True
                        break
                if found_spot: break
            
            if not found_spot:
                rack_num_for_bin_type[bin_type] += 1

    # --- Assign the generated data and rename columns ---
    df_processed = df_sorted.copy()
    df_processed['Rack'] = rack_list
    df_processed['Rack No 1st'] = r1_list
    df_processed['Rack No 2nd'] = r2_list
    df_processed['Level'] = lvl_list
    df_processed['Cell'] = cell_list
    
    part_no_col, desc_col, model_col, station_col, _ = find_required_columns(df)
    rename_dict = {
        part_no_col: 'Part No', desc_col: 'Description',
        model_col: 'Bus Model', station_col: 'Station No'
    }
    final_columns = {k: v for k, v in rename_dict.items() if k is not None}
    df_processed.rename(columns=final_columns, inplace=True)
    
    return df_processed


def create_location_key(row):
    """Create a unique key for grouping by the newly generated location."""
    return '_'.join([
        str(row.get('Bus Model', '')), str(row.get('Station No', '')),
        str(row.get('Rack', '')), str(row.get('Rack No 1st', '')),
        str(row.get('Rack No 2nd', '')), str(row.get('Level', '')),
        str(row.get('Cell', ''))
    ])

def extract_location_values(row):
    """Extract location values from the processed row's columns."""
    return [
        str(row.get('Bus Model', '')), str(row.get('Station No', '')),
        str(row.get('Rack', '')), str(row.get('Rack No 1st', '')),
        str(row.get('Rack No 2nd', '')), str(row.get('Level', '')),
        str(row.get('Cell', ''))
    ]

# --- PDF Generation Functions (No Changes Here) ---

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

            location_data = [['Line Location'] + location_values]
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
            if status_text: st.text(f"Error processing location {location_key}: {e}")
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
            
            location_data = [['Line Location'] + location_values]
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
            if status_text: st.text(f"Error processing location {location_key}: {e}")
            continue

    if progress_bar: progress_bar.progress(100)
    if elements:
        if status_text: status_text.text("Building PDF document...")
        doc.build(elements)
        buffer.seek(0)
        return buffer
    return None

# --- Main Application UI (UPDATED) ---
def main():
    st.title("🏷️ Rack Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )
    st.markdown("---")

    # --- UI Sidebar ---
    st.sidebar.title("📄 Label Options")
    label_type = st.sidebar.selectbox(
        "Choose Label Format:", ["Single Part", "Multiple Parts"],
        help="Single Part: One part per label. Multiple Parts: Up to two parts per label."
    )

    st.sidebar.title("⚙️ Automation Settings")
    rack_input = st.sidebar.text_input(
        "Enter Storage Line Side Infrastructure", "R",
        help="Enter the static value for the 'Rack' field (e.g., TR, R, S)."
    )
    
    uploaded_file = st.file_uploader(
        "Choose an Excel or CSV file", type=['xlsx', 'xls', 'csv'],
        help="Upload your file with Part No, Description, and Container Type."
    )

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith('.csv') else pd.read_excel(uploaded_file)
            st.success(f"✅ File loaded successfully! Found {len(df)} rows.")
            
            with st.expander("📊 File Preview", expanded=False):
                st.dataframe(df.head(3))

            # --- NEW UI for Global and Bin-Specific Settings ---
            level_selections, bin_capacities = {}, {}
            level_dimensions_by_level = {}
            standard_bin_dimensions = {}

            _, _, _, _, container_col = find_required_columns(df)
            
            if container_col:
                unique_bins = get_unique_bins(df, container_col)
                if unique_bins:
                    st.sidebar.markdown("---")
                    st.sidebar.subheader("Global Dimension Settings")
                    st.sidebar.info("These dimensions apply to all bins and levels.")

                    # 1. Ask for Standard Bin Dimensions ONCE
                    with st.sidebar.expander("Enter Standard Bin Dimensions", expanded=True):
                        st.markdown("Dimensions for **all** container types (Bin A, Bin B, etc.)")
                        b_col1, b_col2, b_col3 = st.columns(3)
                        b_l = b_col1.number_input("L", min_value=0.1, value=40.0, step=0.1, key="std_bin_l", help="Standard Length for ALL bins")
                        b_w = b_col2.number_input("W", min_value=0.1, value=40.0, step=0.1, key="std_bin_w", help="Standard Width for ALL bins")
                        b_h = b_col3.number_input("H", min_value=0.1, value=40.0, step=0.1, key="std_bin_h", help="Standard Height for ALL bins")
                        standard_bin_dimensions = {'L': b_l, 'W': b_w, 'H': b_h}

                    # 2. Ask which levels to configure
                    all_possible_levels = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
                    configurable_levels = st.sidebar.multiselect(
                        "Select Levels to Configure",
                        options=all_possible_levels,
                        default=['A', 'B', 'C', 'D']
                    )

                    # 3. Ask for dimensions for each selected level
                    if configurable_levels:
                        with st.sidebar.expander("Enter Dimensions for Each Level", expanded=True):
                            for level in configurable_levels:
                                st.markdown(f"**Dimensions for Level {level}**")
                                l_col1, l_col2, l_col3 = st.columns(3)
                                l_l = l_col1.number_input("L", min_value=0.1, value=100.0, step=0.1, key=f"lvl_l_{level}", help=f"Length of Level {level}")
                                l_w = l_col2.number_input("W", min_value=0.1, value=50.0, step=0.1, key=f"lvl_w_{level}", help=f"Width of Level {level}")
                                l_h = l_col3.number_input("H", min_value=0.1, value=50.0, step=0.1, key=f"lvl_h_{level}", help=f"Height of Level {level}")
                                level_dimensions_by_level[level] = {'L': l_l, 'W': l_w, 'H': l_h}

                    # 4. Ask for Bin-Specific Assignments and Capacity
                    st.sidebar.markdown("---")
                    st.sidebar.subheader("Bin-Specific Settings")
                    for bin_type in unique_bins:
                        with st.sidebar.expander(f"Settings for {bin_type}", expanded=True):
                            level_selections[bin_type] = st.multiselect(
                                "Assignable Levels",
                                options=configurable_levels, # Can only assign to configured levels
                                default=configurable_levels, key=f"assign_{bin_type}",
                                help=f"Which pre-configured levels can {bin_type} be placed on?"
                            )
                            bin_capacities[bin_type] = st.number_input(
                                "Capacity (parts)",
                                min_value=1, value=10, step=1, key=f"cap_{bin_type}",
                                help=f"How many parts fit inside one {bin_type}?"
                            )
                else:
                    st.warning("No values containing 'Bin' found in the 'Container Type' column.")
            else:
                st.error("Could not find a 'Container Type' column. Automation is disabled.")

            # --- PDF Generation ---
            if st.button("🚀 Generate PDF Labels", type="primary"):
                if not all([rack_input, standard_bin_dimensions, level_dimensions_by_level, level_selections]):
                    st.warning("Please configure all dimension and assignment settings in the sidebar.")
                else:
                    progress_bar, status_text = st.progress(0), st.empty()
                    try:
                        status_text.text("Automating line locations based on dimensions...")
                        df_processed = automate_location_assignment(
                            df, rack_input, level_selections, bin_capacities, 
                            standard_bin_dimensions, level_dimensions_by_level, status_text
                        )
                        
                        if df_processed is not None and not df_processed.empty:
                            gen_func = generate_labels_from_excel_v2 if label_type == "Single Part" else generate_labels_from_excel_v1
                            file_suffix = "single_part" if label_type == "Single Part" else "multi_part"
                            
                            pdf_buffer = gen_func(df_processed, progress_bar, status_text)
                            filename = f"{os.path.splitext(uploaded_file.name)[0]}_{file_suffix}_labels.pdf"
                            
                            if pdf_buffer:
                                status_text.text("✅ PDF generated successfully!")
                                st.download_button(label="📥 Download PDF Labels", data=pdf_buffer.getvalue(), file_name=filename, mime="application/pdf")
                                with st.expander("📈 Generation Statistics", expanded=True):
                                    unique_locations = df_processed['location_key'].nunique()
                                    st.metric("Total Parts Processed", len(df_processed))
                                    st.metric("Unique Locations Created", unique_locations)
                                    st.metric("Labels Generated", unique_locations)
                            else:
                                st.error("❌ Failed to generate PDF. Check file data.")
                        else:
                            st.error("❌ Location assignment failed. The processed data is empty.")      
                    except Exception as e:
                        st.error(f"❌ An unexpected error occurred: {e}")
                    finally:
                        progress_bar.empty(); status_text.empty()
        except Exception as e:
            st.error(f"❌ Error reading file: {e}")
    else:
        st.info("👆 Upload a file to begin.")
        with st.expander("📋 Instructions", expanded=True):
            st.markdown("""
            ### How to use this tool:
            1.  **Select Label Format & Enter Rack**: Choose your label style and enter the base **Rack** value (e.g., R).
            2.  **Upload Your File**: It must contain columns for Part Number and Container Type.
            3.  **Configure Global Dimensions**:
                -   Enter the **Standard Bin Dimensions** once. This applies to all bin types in your file.
                -   Select which **Levels** you want to use (A, B, C, etc.).
                -   For each level you selected, enter its unique **Level Dimensions**.
            4.  **Configure Bin-Specific Settings**:
                -   For each bin type, select which of the configured levels it can be placed on.
                -   Enter the part **Capacity** (how many parts fit *inside* one bin).
            5.  **Generate & Download**: Click the generate button to create your PDF.
            """)

if __name__ == "__main__":
    main()
