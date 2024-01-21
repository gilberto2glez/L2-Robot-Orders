from importlib.metadata import files
from typing import Type
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
import os
from zipfile import ZipFile

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    
    open_robot_order_website()
    download_orders_csv()
    close_annoying_modal()
    dic_robot_model_to_part_number = capture_table_from_website()
    complete_orders_with_csv_data(dic_robot_model_to_part_number)
    archive_receipts()
    

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_orders_csv():
    """downloads orders csv from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def complete_orders_with_csv_data(dic_robot_model_to_part_number):
    """Read data from csv and complete orders"""
    csv_table = read_csv_file("orders.csv")
    
    for row in csv_table:
        row["Head"] = dic_robot_model_to_part_number[row["Head"]]
        row["Body"] = dic_robot_model_to_part_number[row["Body"]]
        fill_the_form(row)
        preview_bot()
        file_name_and_location = "output/" + row["Order number"]
        collect_results(file_name_and_location)
        order_bot()
        export_as_pdf(file_name_and_location)
        order_another_bot()
        embed_screenshot_to_receipt(file_name_and_location,file_name_and_location)
    

def read_csv_file(file_path):
    """Reads a csv file and returns a table"""
    tables = Tables()
    table = tables.read_table_from_csv(file_path, header=True)
    return table

def close_annoying_modal():
    """Close modal when opening browser"""
    page = browser.page()
    page.click("xpath=//button[contains(.,'OK')]")

def fill_the_form(order):
    """Fills the order based on the given order data"""
    
    page = browser.page()
    page.select_option("#head",order["Head"] + " head")
    page.click("xpath=//label[contains(.,'" + order["Body"] + " body')]")
    page.fill("xpath=//label[contains(.,'3. Legs:')]/../input",order["Legs"])
    page.fill("#address",order["Address"])


def capture_table_from_website():
    # Initialize the browser and a new webpage
    page = browser.page()
    page.click("xpath=//button[contains(.,'Show model info')]")
    # Navigate to the website containing the table
    page.wait_for_load_state()

    # Locate the table element, you might need to adjust the selector
    table_locator = "xpath=//*[@id='model-info']"
    table_element = page.locator(table_locator)
    table_element.wait_for(timeout=5000, state="visible")

    # Ensure the table is visible
    assert table_element.is_visible()

    # Extract the table data
    # You might need to iterate over rows and cells depending on the table structure
    table_data = table_element.inner_text()

    # Process the table data as needed
    return string_to_dict(table_data)


def string_to_dict(table_string, delimiter='\t'):
    # Split the string into rows
    rows = table_string.strip().split('\n')
    
    # Split each row into columns and create a dictionary
    # Assuming the first column contains values and the second column contains keys
    # Also assuming that the first row is the header which we discard
    table_dict = {row.split(delimiter)[1]: row.split(delimiter)[0] for row in rows[1:]}
    return table_dict

def collect_results(image_name):
    """Take a screenshot of the page"""
    page = browser.page()
    page.wait_for_load_state()
    page.locator("#robot-preview-image").screenshot(path= image_name + ".png")

def export_as_pdf(receipt_name):
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, receipt_name + ".pdf")

def preview_bot():
    """click on the preview button"""
    page = browser.page()
    page.click("#preview")

def order_bot():
    """click on the order button"""
    page = browser.page()
    page.click("#order")
    page.wait_for_load_state()
    page = browser.page()
    if page.is_visible("#order"):
        order_bot()

def order_another_bot():
    """Click on order another bot and removes the modal"""
    page = browser.page()
    page.click("#order-another")
    close_annoying_modal()

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Appends the screenshot to the end of the pdf"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path = screenshot + ".png",
        source_path = pdf_file + ".pdf",
        output_path = pdf_file + ".pdf"
    )

def archive_receipts():
    """archive in receipts pdf all the pdfs in output"""
    zipObj = ZipFile('output/receipts.zip', 'w')
    for file in os.listdir("output"):
        if file.endswith(".pdf"):
            zipObj.write("output/" + file)
    zipObj.close()
