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

# Style definitions
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
    
    # Dynamic font sizing based on description length
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
        # Truncate very long descriptions to prevent overflow
        desc = desc[:100] + "..." if len(desc) > 100 else desc
    
    # Create a custom style for this description
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

def extract_location_values(row, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col):
    """Extract location values from separate Excel columns."""
    location_values = [''] * 7
    
    # Extract values from separate columns
    location_values[0] = str(row.get(bus_model_col, '')) if bus_model_col and bus_model_col in row else ''
    location_values[1] = str(row.get(station_no_col, '')) if station_no_col and station_no_col in row else ''
    location_values[2] = str(row.get(rack_col, '')) if rack_col and rack_col in row else ''
    
    # Handle RACK NO digits - check for separate columns first
    if rack_no_1st_col and rack_no_1st_col in row:
        location_values[3] = str(row.get(rack_no_1st_col, ''))
    elif rack_no_col and rack_no_col in row:
        # Fallback to splitting single RACK NO column if separate columns don't exist
        rack_no_value = str(row.get(rack_no_col, ''))
        if rack_no_value and len(rack_no_value) >= 1:
            location_values[3] = rack_no_value[0]  # 1st digit
        else:
            location_values[3] = ''
    else:
        location_values[3] = ''
    
    if rack_no_2nd_col and rack_no_2nd_col in row:
        location_values[4] = str(row.get(rack_no_2nd_col, ''))
    elif rack_no_col and rack_no_col in row:
        # Fallback to splitting single RACK NO column if separate columns don't exist
        rack_no_value = str(row.get(rack_no_col, ''))
        if rack_no_value and len(rack_no_value) >= 2:
            location_values[4] = rack_no_value[1]  # 2nd digit
        else:
            location_values[4] = ''
    else:
        location_values[4] = ''
    
    location_values[5] = str(row.get(level_col, '')) if level_col and level_col in row else ''
    location_values[6] = str(row.get(cell_col, '')) if cell_col and cell_col in row else ''
    
    return location_values

def find_location_columns(df):
    """Find location-related columns in the DataFrame."""
    cols = [col.upper() for col in df.columns.tolist()]
    
    # Find columns for location components
    bus_model_col = next((col for col in cols if 'BUS' in col and 'MODEL' in col), 
                        next((col for col in cols if 'BUS' in col), None))
    
    station_no_col = next((col for col in cols if 'STATION' in col and ('NO' in col or 'NUM' in col)), 
                         next((col for col in cols if 'STATION' in col), None))
    
    rack_col = next((col for col in cols if 'RACK' in col and 'NO' not in col), None)
    
    # Look for separate 1st and 2nd digit columns first
    rack_no_1st_col = next((col for col in cols if 'RACK' in col and 'NO' in col and ('1ST' in col or '1' in col and 'DIGIT' in col)), None)
    rack_no_2nd_col = next((col for col in cols if 'RACK' in col and 'NO' in col and ('2ND' in col or '2' in col and 'DIGIT' in col)), None)
    
    # Look for general RACK NO column as fallback
    rack_no_col = next((col for col in cols if 'RACK' in col and ('NO' in col or 'NUM' in col) and '1ST' not in col and '2ND' not in col and '1' not in col and '2' not in col), None)
    
    level_col = next((col for col in cols if 'LEVEL' in col), None)
    
    cell_col = next((col for col in cols if 'CELL' in col), None)
    
    return bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col

def find_packaging_factor_column(df):
    """Find the packaging factor column in the DataFrame."""
    cols = [col.upper() for col in df.columns.tolist()]
    
    # Look for packaging factor column
    packaging_factor_col = next((col for col in cols if 'PACKAGING' in col and 'FACTOR' in col), None)
    
    return packaging_factor_col

def determine_label_type(row, packaging_factor_col):
    """Determine label type based on packaging factor value."""
    if packaging_factor_col and packaging_factor_col in row:
        packaging_factor = row[packaging_factor_col]
        
        # Convert to float for comparison
        try:
            factor_value = float(packaging_factor)
            if factor_value == 1.0:
                return "single"
            elif factor_value == 0.5:
                return "multiple"
        except (ValueError, TypeError):
            pass
    
    # Default to single if no packaging factor or invalid value
    return "single"

def create_location_key(row, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col):
    """Create a unique key for grouping by location."""
    values = []
    values.append(str(row.get(bus_model_col, '')) if bus_model_col else '')
    values.append(str(row.get(station_no_col, '')) if station_no_col else '')
    values.append(str(row.get(rack_col, '')) if rack_col else '')
    
    # Handle rack no values
    if rack_no_1st_col:
        values.append(str(row.get(rack_no_1st_col, '')) if rack_no_1st_col else '')
    elif rack_no_col:
        rack_no_value = str(row.get(rack_no_col, ''))
        values.append(rack_no_value[0] if rack_no_value else '')
    else:
        values.append('')
        
    if rack_no_2nd_col:
        values.append(str(row.get(rack_no_2nd_col, '')) if rack_no_2nd_col else '')
    elif rack_no_col:
        rack_no_value = str(row.get(rack_no_col, ''))
        values.append(rack_no_value[1] if len(rack_no_value) > 1 else '')
    else:
        values.append('')
    
    values.append(str(row.get(level_col, '')) if level_col else '')
    values.append(str(row.get(cell_col, '')) if cell_col else '')
    
    return '_'.join(values)

def generate_labels_from_excel_v1(df, progress_bar=None, status_text=None):
    """Generate labels using version 1 formatting (Multiple Parts)."""
    
    # Create a BytesIO buffer to store the PDF
    buffer = io.BytesIO()
    
    # Set up key measurements
    part_no_height = 1.3 * cm
    desc_loc_height = 0.8 * cm

    # Identify column names in the file
    original_cols = df.columns.tolist()
    df.columns = [col.upper() for col in df.columns]
    cols = df.columns.tolist()

    # Find main columns
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                      next((col for col in cols if col in ['PARTNO', 'PART']), None))
    desc_col = next((col for col in cols if 'DESC' in col), None)

    # Find location columns
    bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col = find_location_columns(df)

    if not part_no_col:
        part_no_col = cols[0]
    if not desc_col:
        desc_col = cols[1] if len(cols) > 1 else part_no_col

    if status_text:
        status_text.text(f"Using columns: Part No: {part_no_col}, Description: {desc_col}")
        status_text.text(f"Location columns: Bus Model: {bus_model_col}, Station: {station_no_col}, Rack: {rack_col}, Rack No: {rack_no_col}, Rack No 1st: {rack_no_1st_col}, Rack No 2nd: {rack_no_2nd_col}, Level: {level_col}, Cell: {cell_col}")

    # Create location key for grouping
    df['location_key'] = df.apply(lambda row: create_location_key(row, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col), axis=1)
    
    # Group parts by location
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)

    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    MAX_LABELS_PER_PAGE = 4
    label_count = 0

    for i, (location_key, group) in enumerate(df_grouped):
        try:
            if progress_bar:
                progress_value = int((i / total_locations) * 100)
                progress_bar.progress(progress_value)
            
            if status_text:
                status_text.text(f"Processing location {i+1}/{total_locations}: {location_key}")

            parts = group.head(2)

            if len(parts) < 2:
                if len(parts) == 1:
                    part1 = parts.iloc[0]
                    part2 = parts.iloc[0]
                else:
                    continue
            else:
                part1 = parts.iloc[0]
                part2 = parts.iloc[1]

            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())

            label_count += 1

            part_no_1 = str(part1[part_no_col])
            desc_1 = str(part1[desc_col])
            part_no_2 = str(part2[part_no_col])
            desc_2 = str(part2[desc_col])
            
            # Extract location values from separate columns
            location_values = extract_location_values(part1, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col)

            # Create tables for both parts with dynamic description formatting
            part_table = Table(
                [['Part No', format_part_no_v1(part_no_1)],
                 ['Description', format_description_v1(desc_1)]],
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_loc_height]
            )

            part_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTRE'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),
                ('VALIGN', (0, 1), (0, 1), 'MIDDLE'),
                ('VALIGN', (1, 1), (1, 1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),
            ]))

            part_table2 = Table(
                [['Part No', format_part_no_v1(part_no_2)],
                 ['Description', format_description_v1(desc_2)]],
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_loc_height]
            )

            part_table2.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTRE'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),
                ('VALIGN', (0, 1), (0, 1), 'MIDDLE'),
                ('VALIGN', (1, 1), (1, 1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),
            ]))

            # Location table
            location_data = [['Part Location'] + location_values]
            first_col_width = 4 * cm
            location_widths = [first_col_width]
            remaining_width = 11 * cm
            col_proportions = [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]
            total_proportion = sum(col_proportions)
            adjusted_widths = [w * remaining_width / total_proportion for w in col_proportions]
            location_widths.extend(adjusted_widths)

            location_table = Table(
                location_data,
                colWidths=location_widths,
                rowHeights=desc_loc_height
            )

            location_colors = [
                colors.HexColor('#E9967A'),
                colors.HexColor('#ADD8E6'),
                colors.HexColor('#90EE90'),
                colors.HexColor('#FFD700'),
                colors.HexColor('#ADD8E6'),
                colors.HexColor('#E9967A'),
                colors.HexColor('#90EE90')
            ]

            location_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('VALIGN', (1, 0), (-1, 0), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, 0), 16),
                ('FONTSIZE', (1, 0), (-1, -1), 14),
            ]

            for j, color in enumerate(location_colors):
                location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))

            location_table.setStyle(TableStyle(location_style))

            elements.append(part_table)
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(part_table2)
            elements.append(location_table)
            elements.append(Spacer(1, 0.2 * cm))

        except Exception as e:
            if status_text:
                status_text.text(f"Error processing location {location_key}: {e}")
            continue

    if progress_bar:
        progress_bar.progress(100)

    if elements:
        if status_text:
            status_text.text("Building PDF document...")
        doc.build(elements)
        buffer.seek(0)
        return buffer
    else:
        if status_text:
            status_text.text("No labels were generated. Check if the Excel file has the expected columns.")
        return None

def generate_labels_from_excel_v2(df, progress_bar=None, status_text=None):
    """Generate labels using version 2 formatting (Single Part)."""
    
    buffer = io.BytesIO()
    
    # Set up key measurements
    part_no_height = 1.9 * cm
    desc_height = 2.1 * cm
    loc_height = 0.9 * cm

    # Identify column names
    original_cols = df.columns.tolist()
    df.columns = [col.upper() for col in df.columns]
    cols = df.columns.tolist()

    # Find main columns
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                      next((col for col in cols if col in ['PARTNO', 'PART']), None))
    desc_col = next((col for col in cols if 'DESC' in col), None)

    # Find location columns
    bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col = find_location_columns(df)

    if not part_no_col:
        part_no_col = cols[0]
    if not desc_col:
        desc_col = cols[1] if len(cols) > 1 else part_no_col

    if status_text:
        status_text.text(f"Using columns: Part No: {part_no_col}, Description: {desc_col}")
        status_text.text(f"Location columns: Bus Model: {bus_model_col}, Station: {station_no_col}, Rack: {rack_col}, Rack No: {rack_no_col}, Rack No 1st: {rack_no_1st_col}, Rack No 2nd: {rack_no_2nd_col}, Level: {level_col}, Cell: {cell_col}")

    # Create location key for grouping
    df['location_key'] = df.apply(lambda row: create_location_key(row, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col), axis=1)
    
    # Group parts by location
    df_grouped = df.groupby('location_key')
    total_locations = len(df_grouped)

    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    MAX_LABELS_PER_PAGE = 4
    label_count = 0

    for i, (location_key, group) in enumerate(df_grouped):
        try:
            if progress_bar:
                progress_value = int((i / total_locations) * 100)
                progress_bar.progress(progress_value)
            
            if status_text:
                status_text.text(f"Processing location {i+1}/{total_locations}: {location_key}")

            parts = group.head(2)

            if len(parts) < 2:
                if len(parts) == 1:
                    part1 = parts.iloc[0]
                else:
                    continue
            else:
                part1 = parts.iloc[0]

            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())

            label_count += 1

            part_no = str(part1[part_no_col])
            desc = str(part1[desc_col])
            
            # Extract location values from separate columns
            location_values = extract_location_values(part1, bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col)

            # Part table with enhanced formatting
            part_table = Table(
                [['Part No', format_part_no_v2(part_no)],
                 ['Description', format_description(desc)]],
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_height]
            )

            part_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                ('VALIGN', (1, 0), (1, 0), 'TOP'),
                ('VALIGN', (0, 1), (0, 1), 'MIDDLE'),
                ('VALIGN', (1, 1), (1, 1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (1, 0), (1, 0), 10),
                ('BOTTOMPADDING', (1, 0), (1, 0), 5),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),
            ]))

            # Location table
            location_data = [['Part Location'] + location_values]
            location_widths = [4*cm]
            remaining_width = 11 * cm
            col_widths = [1.7, 2.9, 1.3, 1.2, 1.3, 1.3, 1.3]
            total_proportion = sum(col_widths)
            location_widths.extend([w * remaining_width / total_proportion for w in col_widths])

            location_table = Table(
                location_data,
                colWidths=location_widths,
                rowHeights=loc_height,
            )

            location_colors = [
                colors.HexColor('#E9967A'),
                colors.HexColor('#ADD8E6'),
                colors.HexColor('#90EE90'),
                colors.HexColor('#FFD700'),
                colors.HexColor('#ADD8E6'),
                colors.HexColor('#E9967A'),
                colors.HexColor('#90EE90')
            ]

            location_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('VALIGN', (1, 0), (-1, 0), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, 0), 16),
                ('FONTSIZE', (1, 0), (-1, -1), 16),
            ]

            for j, color in enumerate(location_colors):
                location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))

            location_table.setStyle(TableStyle(location_style))

            elements.append(part_table)
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(location_table)
            elements.append(Spacer(1, 0.2 * cm))

        except Exception as e:
            if status_text:
                status_text.text(f"Error processing location {location_key}: {e}")
            continue

    if progress_bar:
        progress_bar.progress(100)

    if elements:
        if status_text:
            status_text.text("Building PDF document...")
        doc.build(elements)
        buffer.seek(0)
        return buffer
    else:
        if status_text:
            status_text.text("No labels were generated.")
        return None

def generate_labels_automatically(df, progress_bar=None, status_text=None):
    """Generate labels automatically based on packaging factor values."""
    
    # Find packaging factor column
    packaging_factor_col = find_packaging_factor_column(df)
    
    if not packaging_factor_col:
        if status_text:
            status_text.text("No 'Packaging Factor' column found. Using default single part labels.")
        return generate_labels_from_excel_v2(df, progress_bar, status_text), "single"
    
    # Get unique packaging factor values
    df_upper = df.copy()
    df_upper.columns = [col.upper() for col in df_upper.columns]
    
    unique_factors = df_upper[packaging_factor_col].unique()
    
    if status_text:
        status_text.text(f"Found packaging factor values: {unique_factors}")
    
    # Check if we have both single and multiple part types
    has_single = any(float(val) == 1.0 for val in unique_factors if str(val).replace('.', '').isdigit())
    has_multiple = any(float(val) == 0.5 for val in unique_factors if str(val).replace('.', '').isdigit())
    
    if has_single and has_multiple:
        # Mixed packaging factors - need to generate separate PDFs
        if status_text:
            status_text.text("Mixed packaging factors detected. Generating separate label sets.")
        
        # Split dataframe by packaging factor
        df_single = df_upper[df_upper[packaging_factor_col].astype(str).str.replace('.', '').str.isdigit() & 
                            (df_upper[packaging_factor_col].astype(float) == 1.0)]
        df_multiple = df_upper[df_upper[packaging_factor_col].astype(str).str.replace('.', '').str.isdigit() & 
                              (df_upper[packaging_factor_col].astype(float) == 0.5)]
        
        # Generate labels for each type
        single_pdf = generate_labels_from_excel_v2(df_single, progress_bar, status_text) if not df_single.empty else None
        multiple_pdf = generate_labels_from_excel_v1(df_multiple, progress_bar, status_text) if not df_multiple.empty else None
        
        return (single_pdf, multiple_pdf), "mixed"
    
    elif has_multiple:
        # Only multiple part labels
        if status_text:
            status_text.text("Generating multiple part labels (packaging factor 0.5).")
        return generate_labels_from_excel_v1(df, progress_bar, status_text), "multiple"
    
    else:
        # Default to single part labels
        if status_text:
            status_text.text("Generating single part labels (packaging factor 1.0).")
        return generate_labels_from_excel_v2(df, progress_bar, status_text), "single"

# Main Streamlit app
def main():
    st.title("🏷️ Part Label Generator")
    st.write("Upload your Excel file to generate part labels as PDF")
    
    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Read Excel file
            df = pd.read_excel(uploaded_file)
            
            st.success(f"File uploaded successfully! Found {len(df)} rows.")
            
            # Display first few rows
            st.subheader("Data Preview")
            st.dataframe(df.head())
            
            # Label generation options
            st.subheader("Label Generation Options")
            
            generation_mode = st.radio(
                "Select generation mode:",
                ["Automatic (based on packaging factor)", "Manual selection"]
            )
            
            if generation_mode == "Automatic (based on packaging factor)":
                if st.button("Generate Labels Automatically"):
                    with st.spinner("Generating labels..."):
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        result, label_type = generate_labels_automatically(df, progress_bar, status_text)
                        
                        if label_type == "mixed":
                            single_pdf, multiple_pdf = result
                            
                            if single_pdf:
                                st.success("Single part labels generated successfully!")
                                st.download_button(
                                    label="Download Single Part Labels PDF",
                                    data=single_pdf.getvalue(),
                                    file_name="single_part_labels.pdf",
                                    mime="application/pdf"
                                )
                            
                            if multiple_pdf:
                                st.success("Multiple part labels generated successfully!")
                                st.download_button(
                                    label="Download Multiple Part Labels PDF",
                                    data=multiple_pdf.getvalue(),
                                    file_name="multiple_part_labels.pdf",
                                    mime="application/pdf"
                                )
                        
                        elif result:
                            st.success(f"Labels generated successfully! ({label_type} part format)")
                            filename = f"{label_type}_part_labels.pdf"
                            st.download_button(
                                label="Download Labels PDF",
                                data=result.getvalue(),
                                file_name=filename,
                                mime="application/pdf"
                            )
                        else:
                            st.error("Failed to generate labels. Please check your data.")
            
            else:  # Manual selection
                label_format = st.selectbox(
                    "Select label format:",
                    ["Single Part (Version 2)", "Multiple Parts (Version 1)"]
                )
                
                if st.button("Generate Labels"):
                    with st.spinner("Generating labels..."):
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        if label_format == "Single Part (Version 2)":
                            pdf_buffer = generate_labels_from_excel_v2(df, progress_bar, status_text)
                            filename = "single_part_labels.pdf"
                        else:
                            pdf_buffer = generate_labels_from_excel_v1(df, progress_bar, status_text)
                            filename = "multiple_part_labels.pdf"
                        
                        if pdf_buffer:
                            st.success("Labels generated successfully!")
                            st.download_button(
                                label="Download Labels PDF",
                                data=pdf_buffer.getvalue(),
                                file_name=filename,
                                mime="application/pdf"
                            )
                        else:
                            st.error("Failed to generate labels. Please check your data.")
            
            # Data info
            st.subheader("Data Information")
            st.write(f"**Total rows:** {len(df)}")
            st.write(f"**Columns:** {', '.join(df.columns.tolist())}")
            
            # Show column mapping
            with st.expander("Column Mapping Details"):
                df_temp = df.copy()
                df_temp.columns = [col.upper() for col in df_temp.columns]
                
                # Find columns
                bus_model_col, station_no_col, rack_col, rack_no_col, rack_no_1st_col, rack_no_2nd_col, level_col, cell_col = find_location_columns(df_temp)
                packaging_factor_col = find_packaging_factor_column(df_temp)
                
                st.write("**Location Columns Detected:**")
                st.write(f"- Bus Model: {bus_model_col}")
                st.write(f"- Station No: {station_no_col}")
                st.write(f"- Rack: {rack_col}")
                st.write(f"- Rack No: {rack_no_col}")
                st.write(f"- Rack No 1st: {rack_no_1st_col}")
                st.write(f"- Rack No 2nd: {rack_no_2nd_col}")
                st.write(f"- Level: {level_col}")
                st.write(f"- Cell: {cell_col}")
                st.write(f"- Packaging Factor: {packaging_factor_col}")
        
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
            st.write("Please make sure the file is a valid Excel file (.xlsx or .xls)")

if __name__ == "__main__":
    main()
