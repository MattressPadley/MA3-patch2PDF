import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Frame, PageTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import os
import sys

# Default directory for XML files
DEFAULT_XML_DIR = '/Users/mhadley/MALightingTechnology/gma3_library/patch/stages/'

# Function to convert XML data to DataFrame
def xml_to_df(xml_data):
    # Parse XML data
    root = ET.fromstring(xml_data)
    data = []
    
    # Find the Stage element
    stage = root.find('Stage')
    if stage is None:
        print("Error: No Stage element found in XML")
        return pd.DataFrame()
        
    # Find the Fixtures element
    fixtures = stage.find('Fixtures')
    if fixtures is None:
        print("Error: No Fixtures element found in XML")
        return pd.DataFrame()
    
    # Iterate over each 'Fixture' element in the XML data
    for fixture in fixtures.findall('Fixture'):
        # Extract attributes
        name = fixture.get('Name', '')
        mode = fixture.get('Mode', '').split('.')[-1] if fixture.get('Mode') else ''
        fid = fixture.get('FID', '')
        patch = fixture.get('Patch', '')
        
        # Append the attributes to the data list
        data.append([fid, name, mode, patch])
        
    # Convert the data list to a DataFrame
    df = pd.DataFrame(data, columns=['FID', 'Name', 'Mode', 'Patch'])
    # Sort by FID
    df = df.sort_values('FID').reset_index(drop=True)
    return df

class PageNumCanvas:
    def __init__(self, file_path, stage_name):
        self.file_path = file_path
        self.stage_name = stage_name
        
    def on_page(self, canvas, doc, data):
        # Add a title
        title = self.stage_name
        canvas.setFont('Helvetica-Bold', 16)
        canvas.drawString(doc.leftMargin, doc.height + doc.topMargin + 36, title)
        # Add a subheading
        subheading = 'Patchlist'
        canvas.setFont('Helvetica', 14)
        canvas.drawString(doc.leftMargin, doc.height + doc.topMargin + 18, subheading)
        # Add page numbers
        total_pages = (len(data) - 1) // 40 + 1 if data else 1
        page_number = f'{doc.page} / {total_pages}'
        canvas.setFont('Helvetica', 10)
        canvas.drawString(doc.width + doc.leftMargin - 50, doc.bottomMargin - 20, page_number)
        # Add revision date
        revision_date = f'Revision: {datetime.now().strftime("%Y-%m-%d")}'
        canvas.drawString(doc.leftMargin, doc.bottomMargin - 20, revision_date)

def get_xml_files():
    """Get list of XML files from the default directory"""
    if not os.path.exists(DEFAULT_XML_DIR):
        print(f"Error: Directory {DEFAULT_XML_DIR} not found")
        sys.exit(1)
    return [f for f in os.listdir(DEFAULT_XML_DIR) if f.endswith('.xml')]

def convert_to_pdf_and_csv(file_path):
    print(f"Converting {os.path.basename(file_path)} to PDF and CSV...")
    # Read XML data from file
    with open(file_path, 'r') as file:
        xml_data = file.read()
    
    # Parse XML for stage name
    root = ET.fromstring(xml_data)
    stage = root.find('Stage')
    stage_name = stage.get('Name', os.path.basename(file_path).replace('.xml', '')) if stage is not None else os.path.basename(file_path).replace('.xml', '')

    # Convert XML data to DataFrame
    df = xml_to_df(xml_data)

    # Save CSV file
    csv_file = os.path.join(os.getcwd(), os.path.basename(file_path).replace('.xml', '.csv'))
    df.to_csv(csv_file, index=False)
    print(f"CSV created successfully: {csv_file}")

    # Convert DataFrame to list of lists for PDF
    data = df.values.tolist()
    
    # If no data was found
    if not data:
        print("Warning: No fixture data found in the XML file")
        data = [["No fixture data found"]]
    else:
        # Add column names as the first row
        data.insert(0, df.columns.tolist())

    # Create a PDF document in the current directory
    output_file = os.path.join(os.getcwd(), os.path.basename(file_path).replace('.xml', '.pdf'))
    pdf = SimpleDocTemplate(output_file, pagesize=letter)

    # Create page numbering handler
    page_handler = PageNumCanvas(file_path, stage_name)

    # Create a custom frame that includes a title, subheading, and page numbers
    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id='normal')
    template = PageTemplate(id='test', frames=frame, onPage=lambda canvas, doc: page_handler.on_page(canvas, doc, data), pagesize=letter)
    pdf.addPageTemplates([template])

    # Create a list to store the elements
    elements = []

    # Add a table for each page
    for i in range(0, len(data), 40):  # Change 40 to the number of rows you want per page
        # Create a table
        chunk = data[i:i+40]
        if chunk:  # Only create table if we have data
            table = Table(chunk, repeatRows=1)
            # Add styles
            styles = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 14),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0,0), (-1,-1), 1, colors.black)
            ]
            
            # Add alternating row colors
            for j in range(len(chunk)):
                bg_color = colors.lightgrey if j % 2 == 0 else colors.white
                styles.append(('BACKGROUND', (0, j), (-1, j), bg_color))
            
            table.setStyle(TableStyle(styles))
            elements.append(table)
            
            # Add a page break if not the last chunk
            if i + 40 < len(data):
                elements.append(PageBreak())

    # Build the PDF document
    pdf.build(elements)
    print(f"PDF created successfully: {output_file}")

def main():
    # Get list of XML files
    xml_files = get_xml_files()
    
    if not xml_files:
        print("No XML files found in the directory.")
        sys.exit(1)
    
    # Print available files
    print("\nAvailable XML files:")
    for i, file in enumerate(xml_files, 1):
        print(f"{i}. {file}")
    
    # Get user selection
    while True:
        try:
            choice = input("\nEnter the number of the file to convert (or 'q' to quit): ")
            if choice.lower() == 'q':
                sys.exit(0)
            
            file_index = int(choice) - 1
            if 0 <= file_index < len(xml_files):
                selected_file = xml_files[file_index]
                file_path = os.path.join(DEFAULT_XML_DIR, selected_file)
                convert_to_pdf_and_csv(file_path)
                break
            else:
                print("Invalid selection. Please try again.")
        except ValueError:
            print("Please enter a valid number.")

if __name__ == "__main__":
    main()
