import pandas as pd
import xml.etree.ElementTree as ET
import tkinter as tk
import threading
from datetime import datetime
from tkinter import filedialog
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Frame, PageTemplate
from reportlab.lib import colors
import os

# Function to convert XML data to DataFrame
def xml_to_df(xml_data):
    # Parse XML data
    root = ET.fromstring(xml_data)
    data = []
    # Iterate over each 'Fixture' element in the XML data
    for fixture in root.findall('Fixture'):
        # Extract the 'FID', 'Name', 'Mode', and 'Patch' attributes
        FID = fixture.get('FID')
        Name = fixture.get('Name')
        Mode = fixture.get('Mode').split('.')[2]  # Use the second part of the 'Mode' attribute
        Patch = fixture.get('Patch')
        # Append the attributes to the data list
        data.append([FID, Name, Mode, Patch])
    # Convert the data list to a DataFrame
    df = pd.DataFrame(data, columns=['FID', 'Name', 'Mode', 'Patch'])
    return df

# Function to add a title, subheading, and page numbers
def on_page(canvas, doc, data):
    # Add a title
    title = os.path.basename(file_path).replace('.xml', '')  # Remove '.xml' from the title
    canvas.setFont('Helvetica-Bold', 16)
    canvas.drawString(doc.leftMargin, doc.height + doc.topMargin + 36, title)
    # Add a subheading
    subheading = 'Patchlist'
    canvas.setFont('Helvetica', 14)
    canvas.drawString(doc.leftMargin, doc.height + doc.topMargin + 18, subheading)
    # Add page numbers
    page_number = f'{doc.page} / {len(data) // 40 + 1}'
    canvas.setFont('Helvetica', 10)
    canvas.drawString(doc.width + doc.leftMargin - 50, doc.bottomMargin - 20, page_number)
    # Add revision date
    revision_date = f'Revision: {datetime.now().strftime("%Y-%m-%d")}'
    canvas.drawString(doc.leftMargin, doc.bottomMargin - 20, revision_date)


def convert_to_pdf():
    global file_path
    # Read XML data from file
    with open(file_path, 'r') as file:
        xml_data = file.read()

    # Convert XML data to DataFrame
    df = xml_to_df(xml_data)

    # Convert DataFrame to list of lists
    data = df.values.tolist()
    # Add column names as the first row
    data.insert(0, df.columns.tolist())

    # Create a PDF document
    output_file = os.path.splitext(file_path)[0] + '.pdf'  # Name the output file after the original file
    pdf = SimpleDocTemplate(output_file, pagesize=letter)

  # Create a custom frame that includes a title, subheading, and page numbers
    frame = Frame(pdf.leftMargin, pdf.bottomMargin, pdf.width, pdf.height, id='normal')
    template = PageTemplate(id='test', frames=frame, onPage=lambda canvas, doc: on_page(canvas, doc, data), pagesize=letter)
    pdf.addPageTemplates([template])

    # Create a list to store the elements
    elements = []

    # Add a table for each page
    for i in range(0, len(data), 40):  # Change 40 to the number of rows you want per page
        # Create a table
        table = Table(data[i:i+40], repeatRows=1)
        # Add a title
        table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                   ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                   ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                   ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                   ('FONTSIZE', (0, 0), (-1, 0), 14),
                                   ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                   ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
                                   ('BACKGROUND', (0, 2), (-1, -1), colors.white),
                                   ('GRID', (0,0), (-1,-1), 1, colors.black)]))
        # Alternate row colors
        for j in range(len(data[i:i+40])):
            if j % 2 == 0:
                bg_color = colors.lightgrey
            else:
                bg_color = colors.white
            table.setStyle(TableStyle([('BACKGROUND', (0, j), (-1, j), bg_color)]))
        # Add the table to the elements list
        elements.append(table)
        # Add a page break
        if i+40 < len(data):
            elements.append(PageBreak())

    # Build the PDF document
    pdf.build(elements)

    # Display "Success!" message
    success_label.config(text="Success!")
    root.after(2000, lambda: success_label.config(text=""))  # Clear the message after 2 seconds


def open_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[('XML Files', '*.xml')])  # Show the file dialog
    threading.Thread(target=convert_to_pdf).start()

root = tk.Tk()
root.title('GrandMA3 Patch2PDF')

# Set the window size to 16:9 aspect ratio and a bit bigger
root.geometry('400x200')

# Bring the window to the front
root.attributes('-topmost', True)

label = tk.Label(root, text='GrandMA3 Patch2PDF', font=('Helvetica', 16))
label.pack(pady=10)

open_button = tk.Button(root, text='Open File', command=open_file)
open_button.pack(pady=10)

exit_button = tk.Button(root, text='Exit', command=root.quit)
exit_button.pack(pady=10)

# Add a label to display the "Success!" message
success_label = tk.Label(root, text="", font=('Helvetica', 10))
success_label.pack(side='bottom', pady=5)

credit_label = tk.Label(root, text='by Matt Hadley', font=('Helvetica', 10))
credit_label.pack(side='bottom', pady=5)

root.mainloop()
