import dataiku
from datetime import datetime, timedelta, timezone
from dateutil.relativedelta import relativedelta
from dataikuapi.dssclient import DSSClient
from reportlab.lib import colors
from reportlab.lib.colors import red
from reportlab.lib.pagesizes import landscape, A3, A2, A1
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
import io
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import openpyxl
from openpyxl.styles import Font, Alignment
import time


dday = 0

tz = timezone(timedelta(hours=5, minutes=30))

def getMonthYear(days=dday):
    today = datetime.now(tz) - timedelta(days=days)
    file_date = today.strftime('%B_%Y')
    return file_date

def getMonthYear2(days=dday):
    today = datetime.now(tz) - timedelta(days=days)
    file_date = today.strftime('%B %Y')
    return file_date

def getFileDate(day=dday):
    today = datetime.now(tz) - timedelta(days=day)
    file_date = today.strftime('%Y%m%d')
    return file_date

def getCurrentDate(days=0):
    today = datetime.now(tz) - timedelta(days=days)
    file_date = today.strftime('%d')
    return file_date

def getStartDate():
    # Get today's date
    today = datetime.now(tz)

    # Subtract one month
    last_month = today - relativedelta(months=1)

    # Format the date to 'Month Year'
    file_date = last_month.strftime(f'1-%B-%Y')

    return file_date

def getfirstDate(days=dday):
    today = datetime.now(tz) - timedelta(days=days)
    file_date = today.strftime(f'1-%B-%Y')
    return file_date

def getfifteenDate(days=dday):
    today = datetime.now(tz) - timedelta(days=days)
    file_date = today.strftime(f'15-%B-%Y')
    return file_date

def getsixteenDate():
    # Get today's date
    today = datetime.now(tz)

    # Subtract one month
    last_month = today - relativedelta(months=1)

    # Format the date to 'Month Year'
    file_date = last_month.strftime(f'16-%B-%Y')

    return file_date

def getEndDate():
    # Get today's date
    today = datetime.now(tz)

    first_day_of_current_month = today.replace(day=1)
    # Subtract one month and get the last day of that month
    last_month_end = first_day_of_current_month - timedelta(days=1)

    return last_month_end.strftime('%d-%B-%Y')

def secret_key():
    client = dataiku.api_client()
    auth_info = client.get_auth_info(with_secrets=True)

    #list comprehension to find secret['value'] for key = sharepoint_access_key
    client_secret_fixit = next(
       (secret['value'] for secret in auth_info["secrets"] 
        if secret['key'] == 'client_secret_fixit'), 
       None
    )

    #Check if the API key value is None - it is run by new user without api keys
    if client_secret_fixit is None:
        raise ValueError(
           "Error: You need to insert 'sharepoint_access_key' to User Center -> "
           "Profile and Settings -> My Account -> Other Credentials -> "
           "'sharepoint_access_key' as a secret value, and Save it."
       )        
    return client_secret_fixit

def client_id():
    client = dataiku.api_client()
    auth_info = client.get_auth_info(with_secrets=True)

    #list comprehension to find secret['value'] for key = sharepoint_access_key
    client_id_fixit = next(
       (secret['value'] for secret in auth_info["secrets"] 
        if secret['key'] == 'client_id_fixit'), 
       None
    )

    #Check if the API key value is None - it is run by new user without api keys
    if client_id_fixit is None:
        raise ValueError(
           "Error: You need to insert 'sharepoint_access_key' to User Center -> "
           "Profile and Settings -> My Account -> Other Credentials -> "
           "'sharepoint_access_key' as a secret value, and Save it."
       )
    return client_id_fixit

def create_pdf_report(dataframe, columns, employee_name, start_date, end_date, invalid_city, columns2):
    # Create a BytesIO buffer to hold the PDF
    pdf_buffer = io.BytesIO()
    pdf_filename = "FF_Expense_Report.pdf"  # Define the PDF filename

    # Create a PDF document in landscape orientation with A2 size
    pdf = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A2))
    elements = []

    # Create a custom style for the header
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Helvetica-Bold',  # You can change the font if needed
        fontSize=18,  # Set your desired font size here
        spaceAfter=12,  # Space after the header
        alignment=1  # Center alignment
    )

    legend_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Helvetica-Bold',
        fontSize=12,
        spaceAfter=12,
        alignment=TA_LEFT  # Center alignment for header
    )

    disclaimer_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Helvetica-Bold',
        fontSize=16,
        spaceAfter=12,
        alignment=TA_LEFT  # Center alignment for header
    )

    # Add the header
    header_text = f'Expense Report for {employee_name} dated: {start_date} - {end_date}'
    header = Paragraph(header_text, header_style)
    elements.append(header)
    elements.append(Paragraph("<br/><br/>", getSampleStyleSheet()['BodyText']))

     # Add Legends section
    legends_text = "Legends: PCME: Personal Car Mileage Eligibility, PDE: Per Diem Allowance Eligibility, PCM Amount: Personal Car Mileage Amount"
    legends_paragraph = Paragraph(legends_text, legend_style)
    disclaimer_text = "DISCLAIMER: FF employees shall be responsible for validating these expenses against their policies and calls submitted at any point in time before claiming them in concur.<BR/><BR/><BR/>"
    disclaimer_paragraph = Paragraph(disclaimer_text, disclaimer_style)
    #legends_paragraph = Paragraph(legends_text, getSampleStyleSheet()['BodyText'])
    #elements.append(Paragraph("<br/><br/>", getSampleStyleSheet()['BodyText']))  # Add space before legends
    #elements.append(disclaimer_paragraph)
    elements.append(legends_paragraph)


    # Print the columns of the DataFrame for debugging
    print("DataFrame Columns:", invalid_city.columns.tolist())

    # Extract dates from the DataFrame using columns2
    if isinstance(columns2, list) and columns2:
        # Check if the last column in columns2 exists in the DataFrame
        date_column = columns2[-1]  # Adjust this if the date is in a different column
        if date_column in invalid_city.columns:
            dates = invalid_city[date_column].dropna().unique()  # Get unique dates, drop NaN values
            dates = [str(date) for date in dates]  # Convert dates to string format
        else:
            print(f"Warning: Column '{date_column}' not found in DataFrame.")
            dates = []
    else:
        dates = []

    # Print extracted dates for debugging
    print("Extracted Dates:", dates)

    # Filter the DataFrame to include only the specified columns
    filtered_dataframe = dataframe[columns]
    filtered_dataframe2 = invalid_city[columns2]
     # Remove duplicates
    filtered_dataframe2 = filtered_dataframe2.drop_duplicates()


    # Prepare data for the table with wrapped text
    data = [filtered_dataframe.columns.tolist()] + [
        [Paragraph(str(cell), getSampleStyleSheet()['BodyText']) for cell in row]
        for row in filtered_dataframe.values.tolist()
    ]  # Combine header and data with wrapped text

    data2 = [filtered_dataframe2.columns.tolist()] + [
        [Paragraph(str(cell), getSampleStyleSheet()['BodyText']) for cell in row]
        for row in filtered_dataframe2.values.tolist()
    ]  # Combine header and data with wrapped text

    # Calculate the grand totals for the relevant columns
    grand_total_per_diem = filtered_dataframe['Per-Diem'].sum()
    grand_total_pcm_amount = filtered_dataframe['PCM Amount'].sum()
    grand_total_allowance = filtered_dataframe['Total Allowance'].sum()

    # Append the grand total row
    grand_total_row = ['Grand Total'] + [''] * (len(columns) - 4) + [grand_total_per_diem, grand_total_pcm_amount, grand_total_allowance]  # Adjust based on the number of columns
    data.append([Paragraph(str(cell), getSampleStyleSheet()['BodyText']) for cell in grand_total_row])  # Add grand total row

    # Create a table
    table = Table(data)

    # Set margins
    margin = 20
    page_width = landscape(A2)[0] - 2 * margin  # Leave some margin
    num_columns = len(filtered_dataframe.columns)

    # Define fixed column widths based on the maximum length of the header text
    column_widths = [max(120, len(header) * 6) for header in filtered_dataframe.columns]  # Increased base width
    total_width = sum(column_widths)

    # If total width exceeds page width, scale down
    if total_width > page_width:
        scale_factor = page_width / total_width
        column_widths = [width * scale_factor for width in column_widths]

    # Set the column widths
    table._argW = column_widths

    # Add style to the table
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),  # Header background color
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Header text color
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center align all cells
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header font
        ('FONTSIZE', (0, 0), (-1, 0), 12),  # Set font size for header
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Data font
        ('FONTSIZE', (0, 1), (-1, -1), 10),  # Set font size for all data cells
        ('BOTTOMPADDING', (0, 0), (-1, 0), 15),  # Increased padding for header
        ('TOPPADDING', (0, 0), (-1, 0), 15),  # Increased padding for header
        ('BOTTOMPADDING', (0, 1), (-1, -1), 12),  # Increased padding for data rows
        ('TOPPADDING', (0, 1), (-1, -1), 12),  # Increased padding for data rows
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Data background color
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Grid lines
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertical alignment
        ('LEFTPADDING', (0, 0), (-1, -1), 10),  # Increased left padding
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),  # Increased right padding
        ('BACKGROUND', (-1, -1), (-1, -1), colors.lightgrey),  # Background for grand total row
        ('FONTSIZE', (-1, -1), (-1, -1), 14),  # Font size for grand total row
        ('FONTNAME', (-1, -1), (-1, -1), 'Helvetica-Bold'),  # Bold font for grand total row
    ])

    # Set the style for the table
    try:
        table.setStyle(style)
    except Exception as e:
        print(f"Error setting table style: {e}")

    # Set a minimum height for rows to ensure readability
    min_row_height = 40  # Minimum height for each row
    for i in range(len(data)):
        table._argH[i] = min_row_height  # Set row height

    # Add the table to the elements list
    elements.append(table)
    elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
    elements.append(disclaimer_paragraph)

    missing_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Helvetica-Bold',
        fontSize=18,
        spaceAfter=12,
        alignment=TA_LEFT,
        textColor=red
    )


    #elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing

    #missing town text
    # Print extracted dates for debugging without square brackets
    if dates:
        missing_text = f"***The expenses for the following employees for the mentioned dates have not been generated due to system error. The employee shall check with his/her Manager or SFE SPOC <br/> <br/> for any clarifications before proceeding to claim them manually in Concur (if eligible)***"        #missing_text2 = f"Please claim them manually post validation with your manager."
    else:
        missing_text = f""


    missing_paragraph = Paragraph(missing_text, missing_style)
    #missing_paragraph2 = Paragraph(missing_text2, missing_style)


    table2 = Table(data2)
    num_columns2 = len(filtered_dataframe2.columns)
    ################

    # Define fixed column widths based on the maximum length of the header text
    column_widths2 = [max(120, len(header) * 6) for header in filtered_dataframe2.columns]  # Increased base width
    total_width2 = sum(column_widths2)

    # If total width exceeds page width, scale down
    if total_width2 > page_width:
        scale_factor2 = page_width / total_width2
        column_widths2 = [width * scale_factor2 for width in column_widths2]

    # Set the column widths
    table2._argW = column_widths2

    # Add style to the table
    style2 = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),  # Header background color
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Header text color
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Left align all cells
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header font
        ('FONTSIZE', (0, 0), (-1, 0), 12),  # Set font size for header
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Data font
        ('FONTSIZE', (0, 1), (-1, -1), 10),  # Set font size for all data cells
        ('BOTTOMPADDING', (0, 0), (-1, 0), 15),  # Increased padding for header
        ('TOPPADDING', (0, 0), (-1, 0), 15),  # Increased padding for header
        ('BOTTOMPADDING', (0, 1), (-1, -1), 12),  # Increased padding for data rows
        ('TOPPADDING', (0, 1), (-1, -1), 12),  # Increased padding for data rows
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Data background color
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Grid lines
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertical alignment
        ('LEFTPADDING', (0, 0), (-1, -1), 10),  # Increased left padding
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),  # Increased right padding
    ])
    # Set the style for the table
    try:
        table2.setStyle(style2)
    except Exception as e:
        print(f"Error setting table style: {e}")

    # Set the horizontal alignment of the table
    table2.hAlign = 'LEFT'  # Align the table to the left

    # Set a minimum height for rows to ensure readability
    min_row_height = 40  # Minimum height for each row
    for i in range(len(data2)):
        table2._argH[i] = min_row_height  # Set row height

    # Add the table to the elements list

    header = Paragraph(header_text, header_style)
    if dates:
        elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
        elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
        elements.append(missing_paragraph)
        elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
        #elements.append(missing_paragraph2)
        #elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
        # Add a spacer if needed for better positioning
        elements.append(Spacer(1, 12))  # Adjust the height as needed for spacing
        elements.append(table2)
    else:
        missing_text = f""
        #print("Extracted Dates: None")

    ################

    # Build the PDF
    pdf.build(elements)

    pdf_buffer.seek(0)  # Move to the beginning of the BytesIO buffer

    return pdf_buffer


def create_pdf_report2(dataframe, columns, start_date, end_date,  invalid_city, columns2):
    # Create a BytesIO buffer to hold the PDF
    pdf_buffer = io.BytesIO()
    pdf_filename = "FF_Expense_Report.pdf"  # Define the PDF filename

    # Create a PDF document in landscape orientation with A2 size
    pdf = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A2))
    elements = []

     # Create a custom style for the header
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Helvetica-Bold',  # You can change the font if needed
        fontSize=18,  # Set your desired font size here
        spaceAfter=12,  # Space after the header
        alignment=1  # Center alignment
    )

    legend_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Helvetica-Bold',
        fontSize=12,
        spaceAfter=12,
        alignment=TA_LEFT  # Center alignment for header
    )

    disclaimer_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Helvetica-Bold',
        fontSize=16,
        spaceAfter=12,
        alignment=TA_LEFT  # Center alignment for header
    )

     # Add the consolidated header at the top of the report
    consolidated_header_text = f'Consolidated Expense Report for your Reportees dated: {start_date} - {end_date}'
    #header_style = getSampleStyleSheet()['Heading1']  # Use a heading style for the header
    elements.append(Paragraph(consolidated_header_text, header_style))
    elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Add space after the header

    # Add Legends section
    legends_text = "Legends: PCME: Personal Car Mileage Eligibility, PDE: Per Diem Allowance Eligibility, PCM Amount: Personal Car Mileage Amount"
    legends_paragraph = Paragraph(legends_text, legend_style)
    disclaimer_text = "DISCLAIMER: Managers shall be responsible for validation of the expenses submitted by their team members while providing any approval in Concur. <BR/><BR/><BR/>"
    disclaimer_paragraph = Paragraph(disclaimer_text, disclaimer_style)
    #legends_paragraph = Paragraph(legends_text, getSampleStyleSheet()['BodyText'])
    #elements.append(Paragraph("<br/><br/>", getSampleStyleSheet()['BodyText']))  # Add space before legends
    #elements.append(disclaimer_paragraph)
    elements.append(legends_paragraph)

    # Group the DataFrame by Employee Name
    grouped = dataframe.groupby('Name')

    # Prepare styles
    styles = getSampleStyleSheet()
    header_style = styles['Heading1']
    data_style = styles['BodyText']

    for emp_name, group in grouped:
        # Add employee header
        elements.append(Paragraph(f"Employee: {emp_name}", header_style))

        # Prepare data for the table
        # Include 'Per-Diem' only once in the header
        data = [columns[:-3] + ['Per-Diem', 'PCM Amount', 'Total Allowance']]  # Adjust header to include 'Per-Diem' once
        total_allowance = 0  # Initialize total allowance for the employee
        total_per_diem = 0  # Initialize total for Per-Diem
        total_pcm_amount = 0  # Initialize total for PCM Amount

        for _, row in group.iterrows():
            # Calculate totals for each row
            row_per_diem = row['Per-Diem']
            row_pcm_amount = row['PCM Amount']
            row_total_allowance = row_per_diem + row_pcm_amount

            total_per_diem += row_per_diem  # Accumulate total Per-Diem
            total_pcm_amount += row_pcm_amount  # Accumulate total PCM Amount
            total_allowance += row_total_allowance  # Accumulate total allowance

            # Append the row data, ensuring 'Per-Diem' is included only once
            data.append(
                [Paragraph(str(row[col]), data_style) for col in columns[:-3]] +  # Include all columns except the last three
                [Paragraph(str(row_per_diem), data_style),  # Include 'Per-Diem' once
                 Paragraph(str(row_pcm_amount), data_style),
                 Paragraph(str(row_total_allowance), data_style)]
            )

        # Append the grand total row for the employee
        grand_total_row = ['Grand Total'] + [''] * (len(columns) - 4) + [  # Move 'Grand Total' one column to the right
            Paragraph(str(total_per_diem), data_style),
            Paragraph(str(total_pcm_amount), data_style),
            Paragraph(str(total_allowance), data_style)
        ]
        data.append(grand_total_row)

        # Create a table
        table = Table(data)

        # Set margins and column widths
        margin = 20
        page_width = landscape(A2)[0] - 2 * margin
        column_widths = [max(120, len(header) * 6) for header in columns[:-3]] + [100, 100, 100]  # Adjust for Per-Diem, PCM Amount, and Total Allowance
        total_width = sum(column_widths)

        # Scale down if total width exceeds page width
        if total_width > page_width:
            scale_factor = page_width / total_width
            column_widths = [width * scale_factor for width in column_widths]

        # Set the column widths
        table._argW = column_widths

        # Add style to the table
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),  # Header background color
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Header text color
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center align all cells
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header font
            ('FONTSIZE', (0, 0), (-1, 0), 12),  # Set font size for header
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Data font
            ('FONTSIZE', (0, 1), (-1, -1), 10),  # Set font size for all data cells
            ('BOTTOMPADDING', (0, 0), (-1, 0), 15),  # Increased padding for header
            ('TOPPADDING', (0, 0), (-1, 0), 15),  # Increased padding for header
            ('BOTTOMPADDING', (0, 1), (-1, -1), 12),  # Increased padding for data rows
            ('TOPPADDING', (0, 1), (-1, -1), 12),  # Increased padding for data rows
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Data background color
            ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Grid lines
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertical alignment
        ])
        table.setStyle(style)

        # Add the table to the elements list
        elements.append(table)
        elements.append(Paragraph("<br/>", data_style))  # Add space between tables

     # Add Legends section
    #legends_text = "Legends: PCME: Personal Car Mileage Eligibility, PDE: Per Diem Allowance Eligibility, PCM Amount: Personal Car Mileage Amount"
    #legends_paragraph = Paragraph(legends_text, getSampleStyleSheet()['BodyText'])
    elements.append(Paragraph("<br/><br/>", getSampleStyleSheet()['BodyText']))  # Add space before Disclaimer
    #elements.append(legends_paragraph)
    elements.append(disclaimer_paragraph)


        # Extract dates from the DataFrame using columns2
    if isinstance(columns2, list) and columns2:
        # Check if the last column in columns2 exists in the DataFrame
        date_column = columns2[-1]  # Adjust this if the date is in a different column
        if date_column in invalid_city.columns:
            dates = invalid_city[date_column].dropna().unique()  # Get unique dates, drop NaN values
            dates = [str(date) for date in dates]  # Convert dates to string format
        else:
            print(f"Warning: Column '{date_column}' not found in DataFrame.")
            dates = []
    else:
        dates = []

    # Print extracted dates for debugging
    print("Extracted Dates:", dates)

    filtered_dataframe2 = invalid_city[columns2]
    # Remove duplicates
    filtered_dataframe2 = filtered_dataframe2.drop_duplicates()

    #print(filtered_dataframe2) debug

    data2 = [filtered_dataframe2.columns.tolist()] + [
        [Paragraph(str(cell), getSampleStyleSheet()['BodyText']) for cell in row]
        for row in filtered_dataframe2.values.tolist()
    ]  # Combine header and data with wrapped text



    missing_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Helvetica-Bold',
        fontSize=18,
        spaceAfter=12,
        alignment=TA_LEFT,
        textColor=red
    )


    #elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing

    #missing town text
    # Print extracted dates for debugging without square brackets
    if dates:
        missing_text = f"***The expenses for the following employees for the mentioned dates have not been generated due to system error. These employees may reach out to you or SFE SPOC <br/> <br/> for any clarifications before proceeding to claim them manually in Concur (if eligible)***"
    else:
        missing_text = f""


    missing_paragraph = Paragraph(missing_text, missing_style)
    #missing_paragraph2 = Paragraph(missing_text2, missing_style)


    # Set margins
    margin = 20
    page_width = landscape(A2)[0] - 2 * margin  # Leave some margin

    table2 = Table(data2)
    num_columns2 = len(filtered_dataframe2.columns)
    ################

    # Define fixed column widths based on the maximum length of the header text
    column_widths2 = [max(120, len(header) * 6) for header in filtered_dataframe2.columns]  # Increased base width
    total_width2 = sum(column_widths2)

    # If total width exceeds page width, scale down
    if total_width2 > page_width:
        scale_factor2 = page_width / total_width2
        column_widths2 = [width * scale_factor2 for width in column_widths2]

    # Set the column widths
    table2._argW = column_widths2

    # Add style to the table
    style2 = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),  # Header background color
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Header text color
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Left align all cells
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header font
        ('FONTSIZE', (0, 0), (-1, 0), 12),  # Set font size for header
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Data font
        ('FONTSIZE', (0, 1), (-1, -1), 10),  # Set font size for all data cells
        ('BOTTOMPADDING', (0, 0), (-1, 0), 15),  # Increased padding for header
        ('TOPPADDING', (0, 0), (-1, 0), 15),  # Increased padding for header
        ('BOTTOMPADDING', (0, 1), (-1, -1), 12),  # Increased padding for data rows
        ('TOPPADDING', (0, 1), (-1, -1), 12),  # Increased padding for data rows
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Data background color
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Grid lines
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertical alignment
        ('LEFTPADDING', (0, 0), (-1, -1), 10),  # Increased left padding
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),  # Increased right padding
    ])
    # Set the style for the table
    try:
        table2.setStyle(style2)
    except Exception as e:
        print(f"Error setting table style: {e}")

    # Set the horizontal alignment of the table
    table2.hAlign = 'LEFT'  # Align the table to the left

    # Set a minimum height for rows to ensure readability
    min_row_height = 40  # Minimum height for each row
    for i in range(len(data2)):
        table2._argH[i] = min_row_height  # Set row height


    #header = Paragraph(header_text, header_style)
    if dates:
        elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
        elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
        elements.append(missing_paragraph)
        elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
        #elements.append(missing_paragraph2)
        #elements.append(Paragraph("<br/><br/><br/>", getSampleStyleSheet()['BodyText']))  # Empty paragraph for spacing
        # Add a spacer if needed for better positioning
        elements.append(Spacer(1, 12))  # Adjust the height as needed for spacing
        elements.append(table2)
    else:
        missing_text = f""
        #print("Extracted Dates: None")

    # Build the PDF
    pdf.build(elements)

    pdf_buffer.seek(0)  # Move to the beginning of the BytesIO buffer

    return pdf_buffer


def send_email_with_pdf(email, name, file_path, start_date, end_date):
    from_address = 'inffexpense@merck.com'
    to_address = email
    #to_address = ', '.join(email)
    subject = f"FF Expense Report for {start_date} to {end_date}"

    html = f"""<html> <body> <p>Hi {name},</p>
        <p>Please find your expense report for the period: {start_date} and {end_date}. Please verify the details and upload this report as an attachment while submitting in Concur.</p>


    <p>For any queries, Please reach out to your SFE SPOC or Manager.</p>

    <p>Regards,</p>
    <p>SFE Team <br></p>
        <br>
    </body> </html> """

    print('Sending email...')
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(html, "html"))

    # Attach the PDF file
    with open(file_path, 'rb') as file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{file_path.split("/")[-1]}"')
        msg.attach(part)

    smtp_server = 'mailhost.merck.com'  # Your SMTP server
    smtp_port = 25  # Your SMTP port
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.sendmail(from_address, to_address, msg.as_string())
            #server.sendmail(from_address, email, msg.as_string())
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")
        
def send_email_with_pdf2(email, name, file_path, start_date, end_date):
    from_address = 'inffexpense@merck.com'
    to_address = email
    #to_address = ', '.join(email)
    subject = f"FF Expense Report for your team members "

    html = f"""<html> <body> <p>Hi {name},</p>
        <p>Please find your team's expense report for the period: {start_date} and {end_date}. Please use this report for validate expenses submitted by your team in Concur.</p>


    <p>For any queries, Please reach out to your SFE SPOC or Team member.</p>

    <p>Regards,</p>
    <p>SFE Team <br></p>
        <br>
    </body> </html> """

    print('Sending email...')
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(html, "html"))

    # Attach the PDF file
    with open(file_path, 'rb') as file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{file_path.split("/")[-1]}"')
        msg.attach(part)

    smtp_server = 'mailhost.merck.com'  # Your SMTP server
    smtp_port = 25  # Your SMTP port
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.sendmail(from_address, to_address, msg.as_string())
            #server.sendmail(from_address, email, msg.as_string())
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")
        
