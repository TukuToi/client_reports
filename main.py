import csv
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import Image
from reportlab.platypus import Image
from reportlab.platypus import PageBreak
from reportlab.lib import colors
from datetime import datetime, timedelta
import math
import matplotlib.pyplot as plt
import os

##
# Read the CSV
##
def read_csv():
    csv_data = []
    with open('data.csv', 'r') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)
        for row in reader:
            csv_data.append(row)
    return csv_data

##
# Add logo
##
def add_logo_on_every_page(canvas, doc):
    company_logo = Image('logo.png', width=94, height=16)
    # Place the image at the top-left corner of the page
    company_logo.drawOn(canvas, 10, doc.height + 20)

##
# Pagenumber on first page
##
def add_page_number_first_page(canvas):
    page_num = canvas.getPageNumber()
    text = "Page %s" % page_num
    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(canvas._pagesize[0] - 10, 0.1 * inch, text)

##
# Pagenumber other pages
##
def add_page_number_later_pages(canvas):
    page_num = canvas.getPageNumber()
    text = "Page %s" % page_num
    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(canvas._pagesize[0] - 10, 0.1 * inch, text)

##
# Data for first page
##
def first_page(canvas, doc):
    add_logo_on_every_page(canvas, doc)
    add_page_number_first_page(canvas)
    add_page_number_later_pages(canvas)
    
##
# Data for all other pages
##
def other_pages(canvas, doc):
    add_logo_on_every_page(canvas, doc)
    add_page_number_later_pages(canvas)

##
# Data for workload durations
#
def workload_durations(csv_data):
    total_duration = timedelta()
    for row in csv_data:
        duration_str = row[6]  # Assuming the duration is at index 2 (adjust if needed)
        duration_parts = duration_str.split(':')
        hours = int(duration_parts[0])
        minutes = int(duration_parts[1])
        duration = timedelta(hours=hours, minutes=minutes)
        total_duration += duration
    
    return total_duration

##
# Data for total projects
#
def timeframe(csv_data):
    earliest_date = None
    latest_date = None
    earliest_date_formatted = ''
    latest_date_formatted = ''
    for row in csv_data:
        date_string = row[0]
        date = datetime.strptime(date_string, '%d %b %Y at %H:%M')
        
        # Check if the current date is the earliest or latest date
        if earliest_date is None or date < earliest_date:
            earliest_date = date
        if latest_date is None or date > latest_date:
            latest_date = date
    # Format the earliest and latest dates as desired
    earliest_date_formatted = earliest_date.strftime('%d %b %Y')
    latest_date_formatted = latest_date.strftime('%d %b %Y')
    return earliest_date_formatted, latest_date_formatted

##
# Get all projects
#
def consolidated_columns_items(csv_data, col):
    items_counts = {}
    items = [row[col] for row in csv_data]
    for project in items:
        items_counts[project] = items_counts.get(project, 0) + 1
    return items_counts

##
# Build workload summary table
##
def workload_summary_table(centered_style_normal, doc, csv_data):
    unique_days = set()
    for row in csv_data:
        date_string = row[0]
        date = datetime.strptime(date_string, '%d %b %Y at %H:%M')
        day = date.date()
        unique_days.add(day)
    total_days = len(unique_days)
    total_duration = workload_durations(csv_data)
    total_duration_hours = total_duration.total_seconds() / 3600
    average_hours_per_day = total_duration_hours / total_days
    workload_summary_table_data = [
		[
			Paragraph("<b>Total Hours</b>", centered_style_normal),
			Paragraph("<b>Total Days</b>",centered_style_normal),
			Paragraph("<b>Average per Day</b>", centered_style_normal)
		],
		[
			Paragraph(str(math.ceil(total_duration_hours)), centered_style_normal),
			Paragraph(str(total_days), centered_style_normal),
			Paragraph(str(math.ceil(average_hours_per_day)), centered_style_normal)
		]
	]
    workload_summary_table = Table(workload_summary_table_data)
    # Create a container to hold the table and center it
    workload_summary_container = Table([[workload_summary_table]], colWidths=[doc.width * 0.61])
    workload_summary_container.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER')]))
    workload_summary_table.setStyle(TableStyle([
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
    ]))
    return workload_summary_container

##
# Projects count table
##
def projects_count_table(centered_style_normal, doc, csv_data):
    
    total_projects = len(consolidated_columns_items(csv_data, 3))
    projects_count_table_data = [
		[
			Paragraph("<b>Total Projects</b>", centered_style_normal),
		],
		[
			Paragraph(str(total_projects), centered_style_normal),
		]
	]
    projects_count_table = Table(projects_count_table_data)
    projects_count_container = Table([[projects_count_table]], colWidths=[doc.width * 0.61])
    projects_count_container.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER')]))
    projects_count_table.setStyle(TableStyle([
        ('LINEBELOW', (0, 0), (0, 0), 1, colors.black)  # Add bottom border below column name
    ], cells=[(0, 0), (0, 0)]))
    return projects_count_container

##
# Generate Plots
##
def generate_pie_chart(subject, col, title, csv_data):
    locals()[subject + "_counts"] = consolidated_columns_items(csv_data, col)
    plt.figure(figsize=(16, 9))
    plt.pie(locals()[subject + "_counts"].values(), labels=locals()[subject + "_counts"].keys(), autopct='%1.1f%%')
    plt.title(title, fontsize=18, fontweight='bold')
    plt.legend(loc='upper left', bbox_to_anchor=(-0.5, 1))
    path = subject + ".png"
    plt.savefig(path)
    plt.close()
    return path

def generate_column_chart(subject, col, title, csv_data):
    
    for row in csv_data:
        row[col[0]] = datetime.strptime(row[col[0]], "%d %b %Y at %H:%M")
        
    grouped_data = {}
    for row in csv_data:
        date = row[col[0]].date()
        task = row[col[1]]
        time_invested_minutes = float(row[col[2]].split(':')[0]) * 60 + float(row[col[2]].split(':')[1])
        time_invested_hours = time_invested_minutes / 60  # Convert to hours
        if (date, task) in grouped_data:
            grouped_data[(date, task)] += time_invested_hours
        else:
            grouped_data[(date, task)] = time_invested_hours

    dates = [date for (date, _) in grouped_data.keys()]
    time_invested = list(grouped_data.values())
    tasks = [task for (_, task) in grouped_data.keys()]
    
    unique_tasks = list(set(tasks))
    colors = plt.cm.tab10(range(len(unique_tasks)))

    plt.figure(figsize=(16, 9))  # Adjust the figure size as desired
    plt.bar(dates, time_invested, color=[colors[unique_tasks.index(task)] for task in tasks])
    
    plt.xlabel('Timeline')
    plt.ylabel('Hours')
    plt.title(title, fontsize=18, fontweight='bold')

    legend_handles = [plt.Rectangle((0, 0), 1, 1, color=colors[unique_tasks.index(task)]) for task in unique_tasks]
    plt.legend(legend_handles, unique_tasks)

    plt.xticks(rotation=45)

    plt.tight_layout()
    path = subject + ".png"
    plt.savefig(path)
    plt.close()
    return path

def create_pdf():
    
    csv_data = read_csv()
	# Prepare data
    title = 'Client Report'
    client = [row[2] for row in csv_data][0]
    content = []

    earliest_date_formatted, latest_date_formatted = timeframe(csv_data)

    # Create the PDF document
    doc = SimpleDocTemplate("client_report.pdf", pagesize=letter, leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)

    # Define styles
    styles = getSampleStyleSheet()
    centered_style_normal = styles['Normal']
    centered_style_normal.alignment = TA_CENTER
    centered_style_heading = styles['Title']
    centered_style_heading.alignment = TA_CENTER
    left_style_subheading = styles['Heading3']
    left_style_subheading.alignment = TA_CENTER
    
    # Add title, client and daterange
    title_content = Paragraph(title, centered_style_normal)
    client_content = Paragraph(client, centered_style_heading)
    content.append(Spacer(1, 200))
    content.append(title_content)
    content.append(Spacer(1, 10))
    content.append(client_content)
    content.append(Spacer(1, 10))
    content.append(Paragraph(f"From {earliest_date_formatted} to {latest_date_formatted}", centered_style_normal))
    content.append(Spacer(1, 50))
 
    # Add summary data
    workload_summary_container = workload_summary_table(centered_style_normal, doc, csv_data)
    content.append(Spacer(1, 100))  # Add spacing before the table
    content.append(Spacer((doc.width - (doc.width * 0.61)) / 2, 0))  # Add horizontal spacing for center alignment
    content.append(workload_summary_container)
    
	# Add project count data
    projects_count_container = projects_count_table(centered_style_normal, doc, csv_data)
    content.append(Spacer(1, 50))  # Add spacing before the table
    content.append(Spacer((doc.width - (doc.width * 0.61)) / 2, 0))  # Add horizontal spacing for center alignment
    content.append(projects_count_container)

	# Break the page
    content.append(PageBreak())

    content.append(Spacer(1, 33))
    project_path = generate_pie_chart(subject='project', col=3, title='Report by projects', csv_data=csv_data)
    activity_path = generate_pie_chart(subject='activity', col=4, title='Report by activities', csv_data=csv_data)
    content.append(Image(project_path, width=592, height=333, hAlign="RIGHT"))
    content.append(Image(activity_path, width=592, height=333, hAlign="RIGHT"))
    
	# Break the page
    content.append(PageBreak())

    content.append(Spacer(1, 33))
    column_path = generate_column_chart(subject='columns', col=[0,4,6], title='Activities Over Time', csv_data=csv_data)
    content.append(Image(column_path, width=576, height=324, hAlign="RIGHT"))
	
    scvdata = []
    for row in csv_data:
        selected_columns = [row[0], row[3], row[4], row[5], row[6]]
        scvdata.append(selected_columns)
    table = Table(scvdata)
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Header row text color
    ]))
    num_rows = len(scvdata)
    for row in range(1, num_rows):  # Start from the second row
        if row % 2 == 0:
            table.setStyle(TableStyle([('BACKGROUND', (0, row), (-1, row), colors.white)]))
        else:
            table.setStyle(TableStyle([('BACKGROUND', (0, row), (-1, row), colors.HexColor('#EAEAEA'))]))

    content.append(Paragraph('Billabe', left_style_subheading))
    
    content.append(table)

    doc.build(content, onFirstPage=first_page, onLaterPages=other_pages)

    try:
        os.remove(column_path)
        os.remove(activity_path)
        os.remove(project_path)
        os.remove('data.csv')
    except OSError as e:
        print(f"Error deleting file(s): {e}")

    print("PDF created successfully.")

# Fire it up
create_pdf()