from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase.acroform import AcroForm
import os
import pandas as pd

def create_pdf(output_path, logo_path, organization,members):
    # Create a canvas and set the size
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    
    # Start an AcroForm context
    form = AcroForm(c)
    
    # Draw the title at the top left
    c.setFont("Helvetica-Bold", 18)
    c.drawString(inch, height - inch, organization)

    # Draw the logo at the top right
    logo_size = 100
    c.drawImage(logo_path, width - inch - logo_size, height - 2 * inch, width=logo_size, height=logo_size)
    
    # Draw the subtitle
    c.setFont("Helvetica", 14)
    c.drawString(inch, height - inch * 1.5, "Cyber Dawn Course Registration Form")

    # Draw headers for the table
    c.setFont("Helvetica-Bold", 12)
    c.drawString(inch, height - 2 * inch, "Name")
    c.drawString(width / 2, height - 2 * inch, "Course Selection")

    # Add dropdowns for the course selections
    # The options are blank, then Basic Hunting and IR, then Advanced Hunting and IR, then Management
    options_list = ["Select", "Basic Hunting and IR", "Advanced Hunting and IR", "Management","Not Attending"]
    y_position = height - 2.5 * inch
    
    # Draw 10 rows of names and dropdowns
    # The student name is the first and second columns concatenated in the dataframe
    # Start the students from the second row and go to the end
    for i in range(0, len(members)):
        #Only add if the 4th column is not "Red" or "Intel" and the 5th column, if it exists, is not "Intel Analyst"
        if members.iloc[i,3] != "Red" and members.iloc[i,3] != "Intel" and  members.iloc[i,3] != "Green" and not (len(members.columns) > 4 and members.iloc[i,4] == "Intel Analyst"):
            c.setFont("Helvetica", 12)
            c.drawString(inch, y_position, members.iloc[i,0] + " " + members.iloc[i,1])
            form.choice(name=f'Choice{i+1}', tooltip='Select Course', value='Select',
                options=options_list, x=width / 2, y=y_position - 15, width=150, height=20)
            y_position -= 0.4 * inch
        #Check if y_position is less than 1 inch from the bottom of the page
        if y_position < inch:
            #If so, create a new page
            c.showPage()
            #Draw the title at the top left
            c.setFont("Helvetica-Bold", 18)
            c.drawString(inch, height - inch, organization)
            #Draw the logo at the top right
            c.drawImage(logo_path, width - inch - logo_size, height - 2 * inch, width=logo_size, height=logo_size)
            #Draw the subtitle
            c.setFont("Helvetica", 14)
            c.drawString(inch, height - inch * 1.5, "Cyber Dawn Course Registration Form")
            #Draw headers for the table
            c.setFont("Helvetica-Bold", 12)
            c.drawString(inch, height - 2 * inch, "Name")
            c.drawString(width / 2, height - 2 * inch, "Course Selection")
            y_position = height - 2.5 * inch

    # Finalize and save the PDF
    c.save()

def buildforms():
    #List all files in the ../reference-files/cd-training-form directory that end with .csv
    files = [f for f in os.listdir('../reference-files/cd-training-form') if f.endswith('.csv')]

    for file in files:
        #Turn the csv file into a pandas dataframe
        df = pd.read_csv(f'../reference-files/cd-training-form/{file}')
        #Create a PDF file with the same name as the csv file
        output_file = f'../reference-files/cd-training-form/forms/{file.replace(".csv", ".pdf")}'
        #Create the PDF file
        #The organization name is the name of the csv file
        create_pdf(output_file, "logo.png", file.replace(".csv", ""),df)


if __name__ == '__main__':
    buildforms()