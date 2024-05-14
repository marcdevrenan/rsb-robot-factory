from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

import logging

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=1000
    )
    orders = get_orders()
    open_robot_order_website()
    for order in orders:
        close_annoying_modal()
        fill_and_submit_robot_orders(order)
    archive_receipts()

def open_robot_order_website():
    """Navigates to the webstite"""
    browser.goto('https://robotsparebinindustries.com/#/robot-order')
    
def close_annoying_modal():
    """Closes constitutional rights modal"""
    page = browser.page()
    page.click('button:text("Ok")')

def get_orders():
    """Downloads excel file with customer orders"""
    http = HTTP()
    http.download(url='https://robotsparebinindustries.com/orders.csv', overwrite=True)
    
    table = Tables()
    return table.read_table_from_csv('orders.csv', header=True)

def fill_and_submit_robot_orders(order: str):
    """Order robots based on instructions"""
    page = browser.page()
    
    page.select_option('#head', str(order['Head']))
    page.set_checked(f'#id-body-{str(order["Body"])}', str(order['Body']))
    page.fill('//input[@type="number"]', order['Legs'])
    page.fill('#address', order['Address'])
    
    while True:
        page.click('button:text("Order")')
        
        if page.is_visible('div.alert.alert-danger'):
            logging.error(f'Unexpected error with order {order['Order number']}, trying again...')
        else:
            break
        
    store_receipt_as_pdf(order['Order number'])
    screenshot_robot(order['Order number'])
    embed_screenshot_to_receipt(order['Order number'])
    page.click('button:text("Order another robot")')
    
def store_receipt_as_pdf(order_number: str):
    """Store the order receipt as a PDF file"""
    page = browser.page()
    robot_receipt_html = page.locator('#receipt').inner_html()
    
    pdf = PDF()
    pdf.html_to_pdf(robot_receipt_html, f'output/receipts/robot_receipt_{order_number}.pdf')

def screenshot_robot(order_number: str):
    """Take a screenshot of the robot"""
    page = browser.page()
    page.screenshot(path=f'output/receipts/robot_screenshot_{order_number}.png')

def embed_screenshot_to_receipt(order_number: str):
    """Embed the robot screenshot to the receipt PDF file"""
    pdf = PDF()
    
    pdf.add_files_to_pdf(
        files=[f'output/receipts/robot_receipt_{order_number}.pdf', f'output/receipts/robot_screenshot_{order_number}.png'],
        target_document=f'output/receipts/merged_pdf_{order_number}.pdf'
    )
    
def archive_receipts():
    """Create a ZIP file of receipt PDF files"""
    zip_file = Archive()
    zip_file.archive_folder_with_zip('output/receipts', 'receipts.zip')
