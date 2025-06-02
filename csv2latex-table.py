import csv
import os
import sys
import chardet

def detect_encoding(file_path):
    """Detect file encoding using chardet"""
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read(10000)
            result = chardet.detect(raw_data)
            return result['encoding'] or 'utf-8'
    except Exception:
        return 'utf-8'

def escape_latex(s):
    """Escape special LaTeX characters and handle missing values"""
    if not isinstance(s, str):
        s = str(s)
    
    # Handle missing values
    if s.strip() in ['', 'NA', 'N/A', 'NaN', 'None', '\\', '//', '---', '_']:
        return r'\textemdash'
    
    # Preserve meaningful slashes
    if s.strip() == '/':
        return '/'
    
    # Escape special characters
    replacements = [
        ('\\', r'\textbackslash{}'),
        ('&', r'\&'),
        ('%', r'\%'),
        ('$', r'\$'),
        ('#', r'\#'),
        ('_', r'\_'),
        ('{', r'\{'),
        ('}', r'\}'),
        ('~', r'\textasciitilde{}'),
        ('^', r'\textasciicircum{}'),
        ('<', r'\textless{}'),
        ('>', r'\textgreater{}'),
        ('[', r'{[}'),
        (']', r'{]}'),
    ]
    for orig, repl in replacements:
        s = s.replace(orig, repl)
    
    return s

def process_header_row(row, prev_row=None):
    """
    Convert header row to LaTeX format with partial horizontal lines
    - row: Current header row
    - prev_row: Previous header row (for detecting subgroup boundaries)
    """
    processed = []
    n = len(row)
    
    # Track subgroup boundaries
    subgroup_ranges = []
    start_idx = None
    
    # Find subgroups (cells with content after empty cells)
    for i, cell in enumerate(row):
        if cell and cell != r'\textemdash':
            if start_idx is not None:
                subgroup_ranges.append((start_idx, i-1))
                start_idx = None
        else:
            if start_idx is None:
                start_idx = i
    if start_idx is not None:
        subgroup_ranges.append((start_idx, n-1))
    
    # Process each cell
    for i, cell in enumerate(row):
        if cell and cell != r'\textemdash':
            # Check if this cell is part of a subgroup
            in_subgroup = any(start <= i <= end for start, end in subgroup_ranges)
            
            # Check if previous row has content in this column
            prev_has_content = prev_row and i < len(prev_row) and prev_row[i] and prev_row[i] != r'\textemdash'
            
            # Add partial horizontal line above if needed
            if in_subgroup and not prev_has_content:
                # Find subgroup range for this cell
                subgroup_range = next((r for r in subgroup_ranges if r[0] <= i <= r[1]), None)
                if subgroup_range:
                    start, end = subgroup_range
                    processed.append(r"\cmidrule(lr){" + f"{start+1}-{end+1}" + r"}" + "\n")
            
            processed.append(cell)
        else:
            processed.append('')
    
    return processed

def csv_to_latex_table(input_csv, output_tex, caption, label, 
                       landscape=True, header_lines=1):
    """
    Convert CSV to LaTeX table with multi-line headers and partial horizontal lines
    
    Args:
        input_csv: Path to input CSV file
        output_tex: Path to output .tex file
        caption: Table caption
        label: Table label
        landscape: Use landscape orientation
        header_lines: Number of header lines
    """
    # Detect and handle encoding
    encoding = detect_encoding(input_csv)
    print(f"Detected encoding: {encoding}")
    
    # Read CSV with fallback encodings
    data = []
    encodings_to_try = [encoding, 'latin1', 'cp1252', 'iso-8859-1', 'utf-16']
    for enc in encodings_to_try:
        try:
            with open(input_csv, 'r', encoding=enc) as f:
                reader = csv.reader(f)
                data = list(reader)
                if data: 
                    print(f"Successfully read with {enc} encoding")
                    break
        except Exception as e:
            print(f"Failed with {enc}: {e}")
    
    if not data:
        print("Error: Could not read CSV data")
        return False
    
    # Validate structure
    if len(data) < header_lines:
        print(f"Error: Not enough rows for {header_lines} header lines")
        return False
    
    # Determine number of columns
    ncols = max(len(row) for row in data)
    print(f"Detected {ncols} columns in CSV")
    
    # Pad all rows to ncols
    padded_data = []
    for row in data:
        if len(row) < ncols:
            padded_data.append(row + [''] * (ncols - len(row)))
        else:
            padded_data.append(row)
    data = padded_data
    
    # Process headers
    header_rows = []
    for i in range(header_lines):
        escaped_row = [escape_latex(cell) for cell in data[i]]
        prev_row = header_rows[-1] if header_rows else None
        header_rows.append(process_header_row(escaped_row, prev_row))
    
    # Process data rows
    rows = []
    for i in range(header_lines, len(data)):
        rows.append([escape_latex(cell) for cell in data[i]])
    
    # Generate column widths dynamically
    base_width = 2.0  # Base width in cm
    widths = [base_width] * ncols
    
    # Generate LaTeX
    try:
        with open(output_tex, 'w', encoding='utf-8') as f:
            # Landscape mode
            if landscape:
                f.write(r"\begin{landscape}" + "\n")
            
            f.write(r"{\footnotesize" + "\n")
            f.write(r"\setlength\LTleft{0pt}" + "\n")
            f.write(r"\setlength\LTright{0pt}" + "\n\n")
            
            # Begin longtable
            f.write(r"\begin{longtable}[p]{" + "\n")
            
            # Column specifications
            col_specs = [f">{{\\raggedright\\arraybackslash}}p{{{w}cm}}" for w in widths]
            f.write("    " + "\n    ".join(col_specs) + "\n")
            f.write(r"}" + "\n\n")
            
            # Caption and label
            f.write(r"\caption{" + caption + r"}" + "\n")
            f.write(r"\label{" + label + r"}" + r" \\" + "\n")
            f.write(r"\toprule" + "\n")
            
            # Header rows with partial horizontal lines
            for i, row in enumerate(header_rows):
                # Write the actual header content
                f.write(" & ".join(row) + r" \\" + "\n")
                
                # REMOVED: Midrules between header rows
                # Only add bottom rule after last header row
                if i == len(header_rows) - 1:
                    f.write(r"\bottomrule" + "\n")
            
            f.write(r"\endfirsthead" + "\n\n")
            
            # Continued table header
            f.write(r"\multicolumn{" + str(ncols) + r"}{c}{{\bfseries \tablename\ \thetable{} -- continued from previous page}} \\" + "\n")
            f.write(r"\toprule" + "\n")
            
            for i, row in enumerate(header_rows):
                f.write(" & ".join(row) + r" \\" + "\n")
                if i == len(header_rows) - 1:
                    f.write(r"\bottomrule" + "\n")
            
            f.write(r"\endhead" + "\n\n")
            
            # Table footer
            f.write(r"\bottomrule" + "\n")
            f.write(r"\multicolumn{" + str(ncols) + r"}{r}{{Continued on next page}} \\" + "\n")
            f.write(r"\endfoot" + "\n\n")
            f.write(r"\bottomrule" + "\n")
            f.write(r"\endlastfoot" + "\n\n")
            
            # Data rows
            for row in rows:
                f.write(" & ".join(row) + r" \\" + "\n")
            
            # End of table
            f.write(r"\end{longtable}" + "\n")
            f.write(r"}" + "\n")
            
            if landscape:
                f.write(r"\end{landscape}" + "\n")
        
        return True
    
    except Exception as e:
        print(f"Error writing LaTeX: {e}")
        return False

# ======================================================
# CONFIGURATION SECTION - EDIT THESE VALUES AS NEEDED
# ======================================================

INPUT_CSV = "Book1.csv"          # Path to input CSV file
OUTPUT_TEX = "output.tex"        # Path for the generated LaTeX file
TABLE_CAPTION = "Controller Configurations Summary"
TABLE_LABEL = "tab:controller_configs"
LANDSCAPE_MODE = True            # Set to False for portrait tables
HEADER_LINES = 2                 # Number of header lines

# ======================================================
# MAIN EXECUTION - NO NEED TO EDIT BELOW THIS LINE
# ======================================================

if __name__ == "__main__":
    print(f"Converting CSV to LaTeX table:\n"
          f"  Input:        {INPUT_CSV}\n"
          f"  Output:       {OUTPUT_TEX}\n"
          f"  Landscape:    {'Yes' if LANDSCAPE_MODE else 'No'}\n"
          f"  Header lines: {HEADER_LINES}")
    
    if not os.path.exists(INPUT_CSV):
        print(f"Error: Input file not found at '{INPUT_CSV}'")
        sys.exit(1)
    
    success = csv_to_latex_table(
        input_csv=INPUT_CSV,
        output_tex=OUTPUT_TEX,
        caption=TABLE_CAPTION,
        label=TABLE_LABEL,
        landscape=LANDSCAPE_MODE,
        header_lines=HEADER_LINES
    )
    
    if success:
        print("\nSuccessfully generated LaTeX table!")
        print("Required LaTeX packages:\n"
              r"\usepackage{array, longtable, lscape, booktabs}")
        print("\nExample of how the header will look:")
        print(r"Ref & Methods & data requirements & \multicolumn{3}{c}{Predicted variables} \\")
        print(r"& & & a & b & c \\")
    else:
        print("\nConversion failed. Check error messages above.")