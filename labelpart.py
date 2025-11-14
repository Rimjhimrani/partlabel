import streamlit as st
import pandas as pd
import io
import math
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# --- Page Configuration ---
st.set_page_config(
    page_title="Smart Rack Label Generator",
    page_icon="🏷️",
    layout="wide"
)

# --- Style Definitions (Unchanged) ---
bold_style_v1 = ParagraphStyle(name='Bold_v1', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, leading=20, spaceBefore=2, spaceAfter=2)
bold_style_v2 = ParagraphStyle(name='Bold_v2', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, leading=12, spaceBefore=0, spaceAfter=15)
desc_style = ParagraphStyle(name='Description', fontName='Helvetica', fontSize=20, alignment=TA_LEFT, leading=16, spaceBefore=2, spaceAfter=2)

# --- Formatting Functions (Unchanged) ---
def format_part_no_v1(part_no):
    if not part_no or not isinstance(part_no, str): part_no = str(part_no)
    if len(part_no) > 5:
        return Paragraph(f"<b><font size=17>{part_no[:-5]}</font><font size=22>{part_no[-5:]}</font></b>", bold_style_v1)
    return Paragraph(f"<b><font size=17>{part_no}</font></b>", bold_style_v1)

def format_part_no_v2(part_no):
    if not part_no or not isinstance(part_no, str): part_no = str(part_no)
    if len(part_no) > 5:
        return Paragraph(f"<b><font size=34>{part_no[:-5]}</font><font size=40>{part_no[-5:]}</font></b><br/><br/>", bold_style_v2)
    return Paragraph(f"<b><font size=34>{part_no}</font></b><br/><br/>", bold_style_v2)

def format_description_v1(desc):
    desc = str(desc) if desc else ""
    length = len(desc)
    size = 15 if length <= 30 else 13 if length <= 50 else 11 if length <= 70 else 9
    style = ParagraphStyle(name='Desc_v1', fontName='Helvetica', fontSize=size, leading=size+2)
    return Paragraph(desc[:100] + "..." if length > 100 else desc, style)

def format_description(desc):
    return Paragraph(str(desc) if desc else "", desc_style)

# --- HELPER FUNCTIONS ---

def find_required_columns(df):
    """Find essential columns in the DataFrame."""
    cols = {col.upper(): col for col in df.columns}
    
    part_no = next((v for k, v in cols.items() if 'PART' in k and ('NO' in k or 'NUM' in k or '#' in k)), None)
    desc = next((v for k, v in cols.items() if 'DESC' in k), None)
    model = next((v for k, v in cols.items() if 'MODEL' in k), None)
    station = next((v for k, v in cols.items() if 'STATION' in k), None)
    container = next((v for k, v in cols.items() if 'CONTAINER' in k or 'BIN' in k), None)

    return part_no, desc, model, station, container

def get_unique_bins(df, container_col):
    """Extracts and sorts unique bin types (e.g., Bin A, Bin B)."""
    if not container_col or container_col not in df.columns:
        return []
    # Filter for values containing "Bin" (case insensitive) and sort them
    unique = df[container_col].dropna().astype(str).unique()
    # Sort alphabetically so Bin A comes before Bin B
    return sorted([u for u in unique if 'BIN' in u.upper()])

def get_level_char(index):
    """Converts 0 -> A, 1 -> B, etc."""
    # Simple mapping for A-Z
    if 0 <= index < 26:
        return chr(65 + index)
    return "?" # Fallback if levels exceed Z

# --- CORE LOGIC: PROCESS ASSIGNMENT ---

def process_and_assign_locations(df, rack_input, bin_capacities, status_text=None):
    """
    Assigns locations based on Capacity Roundup logic.
    1. Sorts Bins (Bin A -> Rack 01, Bin B -> Rack 02).
    2. Fills Level A up to Capacity, then moves to Level B.
    """
    part_col, desc_col, mod_col, stat_col, cont_col = find_required_columns(df)
    
    if not part_col or not cont_col:
        st.error("Missing 'Part Number' or 'Container Type' columns.")
        return None

    # Initialize new columns
    df['Rack'] = rack_input
    df['Rack No 1st'] = ''
    df['Rack No 2nd'] = ''
    df['Level'] = ''
    df['Cell'] = '' # Used for calculating position in level

    sorted_bins = get_unique_bins(df, cont_col)
    
    # Create a list to store processed chunks
    processed_chunks = []

    # Process data separate from the detected bins (items that don't match 'Bin')
    # non_bin_df = df[~df[cont_col].astype(str).str.contains('BIN', case=False, na=False)].copy()
    # if not non_bin_df.empty:
    #    processed_chunks.append(non_bin_df)

    for index, bin_type in enumerate(sorted_bins):
        # 1. Identify Rack Number based on Alphabetical Order (A=1, B=2)
        rack_num = index + 1
        rack_str = f"{rack_num:02d}" # 01, 02
        
        # 2. Get user defined capacity for this bin type
        capacity = bin_capacities.get(bin_type, 10) # Default to 10 if error
        if capacity < 1: capacity = 1

        if status_text:
            status_text.text(f"Processing {bin_type}: Rack {rack_str}, Capacity {capacity}/level...")

        # 3. Filter data for this specific bin
        bin_mask = df[cont_col].astype(str) == bin_type
        sub_df = df[bin_mask].copy()
        
        # 4. Apply Logic: Level assignment based on Capacity
        # Reset index to count items 0, 1, 2...
        sub_df = sub_df.reset_index(drop=True)
        
        # Vectorized calculation or simple loop
        for i in range(len(sub_df)):
            # Roundup Logic:
            # If capacity is 5.
            # Index 0-4 -> Level Index 0 (A)
            # Index 5-9 -> Level Index 1 (B)
            level_index = i // capacity
            cell_num = (i % capacity) + 1  # 1 to Capacity
            
            level_char = get_level_char(level_index)
            
            sub_df.at[i, 'Rack No 1st'] = rack_str[0]
            sub_df.at[i, 'Rack No 2nd'] = rack_str[1]
            sub_df.at[i, 'Level'] = level_char
            sub_df.at[i, 'Cell'] = str(cell_num) # Optional: record the position 
            
        processed_chunks.append(sub_df)

    if not processed_chunks:
        return df # Return original if nothing processed
        
    # Recombine all processed parts
    final_df = pd.concat(processed_chunks, ignore_index=True)
    
    # Fill missing standard columns if they were empty
    final_df['Rack'] = rack_input
    
    return final_df

def create_location_key(row):
    return '_'.join([str(row.get('Rack No 1st', '')), str(row.get('Rack No 2nd', '')), str(row.get('Level', ''))])

def extract_location_values(row):
    return [
        str(row.get('Bus Model', '')),
        str(row.get('Station No', '')),
        str(row.get('Rack', '')),
        str(row.get('Rack No 1st', '')),
        str(row.get('Rack No 2nd', '')),
        str(row.get('Level', '')),
        str(row.get('Cell', ''))
    ]

# --- PDF GENERATION (Simplified for Brevity - Logic Same as before) ---
def generate_pdf(df, label_type, progress_bar):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    
    # Group by location (Rack+Level) or just list sequentially? 
    # Usually sequentially based on the sorted input is best for this logic.
    # However, to use the "Multi-Part" feature, we group by key.
    
    df['key'] = df.apply(create_location_key, axis=1)
    
    # If Single Part, we treat every row individually. 
    # If Multi Part, we try to group 2 parts per location IF they share the same location.
    # But based on the logic (Capacity), every part implies a physical bin spot. 
    # Usually, "Bin A" means the box itself. 
    # I will assume standard sequential generation.
    
    # Standardizing column names for PDF function
    part_col, desc_col, _, _, _ = find_required_columns(df)
    
    count = 0
    total = len(df)
    
    # Loop rows
    for i in range(0, total, 2 if label_type == "Multiple Parts" else 1):
        if progress_bar: progress_bar.progress(int((i / total) * 100))
        
        rows = [df.iloc[i]]
        if label_type == "Multiple Parts" and i + 1 < total:
            # Only group if they are in the same rack/level/bin type?
            # For now, we just group sequentially as requested.
            rows.append(df.iloc[i+1])

        if count > 0 and count % 4 == 0: elements.append(PageBreak())
        count += 1

        # Render Labels
        for row in rows:
            p_no = str(row.get(part_col, ''))
            desc = str(row.get(desc_col, ''))
            locs = extract_location_values(row)
            
            # Table 1: Info
            if label_type == "Single Part":
                p_para = format_part_no_v2(p_no)
                d_para = format_description(desc)
                h_part, h_desc = 1.9*cm, 2.1*cm
            else:
                p_para = format_part_no_v1(p_no)
                d_para = format_description_v1(desc)
                h_part, h_desc = 1.3*cm, 0.8*cm

            t1 = Table([['Part No', p_para], ['Description', d_para]], colWidths=[4*cm, 11*cm], rowHeights=[h_part, h_desc])
            t1.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
            
            # Table 2: Location
            # Headers: Model, Station, Rack, R1, R2, Lev, Cell
            data = [['Line Location'] + locs]
            # Widths logic
            col_props = [1.7, 2.9, 1.3, 1.2, 1.3, 1.3, 1.3]
            widths = [4*cm] + [w * (11*cm) / sum(col_props) for w in col_props]
            
            t2 = Table(data, colWidths=widths, rowHeights=0.9*cm)
            bg_cols = [colors.HexColor('#E9967A'), colors.HexColor('#ADD8E6'), colors.HexColor('#90EE90'), 
                       colors.HexColor('#FFD700'), colors.HexColor('#ADD8E6'), colors.HexColor('#E9967A'), colors.HexColor('#90EE90')]
            style_cmds = [('GRID', (0,0), (-1,-1), 1, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]
            for idx, color in enumerate(bg_cols):
                style_cmds.append(('BACKGROUND', (idx+1, 0), (idx+1, 0), color))
            t2.setStyle(TableStyle(style_cmds))
            
            elements.extend([t1, Spacer(1, 0.2*cm), t2, Spacer(1, 0.3*cm)])

    if progress_bar: progress_bar.progress(100)
    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- MAIN APPLICATION ---

def main():
    st.title("🏷️ Smart Rack Label Generator")
    st.markdown("Authomatically calculates Levels based on Bin Capacity.")
    st.markdown("---")

    st.sidebar.header("1. Static Settings")
    rack_input = st.sidebar.text_input("Enter Rack Name (e.g. TR)", "TR")
    label_type = st.sidebar.selectbox("Label Format", ["Single Part", "Multiple Parts"])

    st.sidebar.header("2. Upload Data")
    uploaded_file = st.sidebar.file_uploader("Upload Excel/CSV", type=['xlsx', 'csv'])

    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            _, _, _, _, container_col = find_required_columns(df)

            if container_col:
                # --- DYNAMIC CAPACITY INPUTS ---
                unique_bins = get_unique_bins(df, container_col)
                
                if unique_bins:
                    st.info(f"📂 File Loaded. Detected {len(unique_bins)} bin types.")
                    st.subheader("3. Configure Bin Capacities")
                    
                    # Create a form or columns for inputs
                    bin_capacities = {}
                    cols = st.columns(len(unique_bins) if len(unique_bins) < 4 else 3)
                    
                    for i, bin_name in enumerate(unique_bins):
                        with cols[i % 3]:
                            # Ask user for capacity. 
                            # Prompt: "how many Bin A will go in A level"
                            cap = st.number_input(
                                f"Capacity for {bin_name}", 
                                min_value=1, 
                                value=10, 
                                help=f"How many {bin_name} fit on one Level?"
                            )
                            bin_capacities[bin_name] = cap
                    
                    st.write("---")
                    
                    if st.button("🚀 Generate Labels"):
                        status_text = st.empty()
                        progress_bar = st.progress(0)
                        
                        # Run Logic
                        df_processed = process_and_assign_locations(df, rack_input, bin_capacities, status_text)
                        
                        if df_processed is not None:
                            # Preview Calculation
                            with st.expander("👀 Preview Calculated Locations"):
                                st.dataframe(df_processed[['Part No', 'Container Type', 'Rack No 1st', 'Rack No 2nd', 'Level', 'Cell']].head(10))
                            
                            # Generate PDF
                            status_text.text("Generating PDF...")
                            pdf_data = generate_pdf(df_processed, label_type, progress_bar)
                            
                            st.success("Done!")
                            st.download_button(
                                label="📥 Download Labels PDF",
                                data=pdf_data,
                                file_name="rack_labels.pdf",
                                mime="application/pdf"
                            )
                        
                else:
                    st.warning("No 'Bin' detected in Container Column.")
            else:
                st.error("Could not find 'Container Type' column.")

        except Exception as e:
            st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
