from robocorp.tasks import task
from robocorp import browser
from RPA.FileSystem import FileSystem
from RPA.Tables import Tables

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Archive import Archive

@task

def order_robots_from_RobotSpareBin():
    """Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images."""
    browser.configure(
    slowmo=100,)
        

    download_csv()
    orders = get_orders()
    open_robot_order_website()
    fill_form_with_csv_data (orders)
    archive_receipts()


def download_csv():
    """Download needed .csv"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Reads Orders .csv"""
    lib = Tables()
    return lib.read_table_from_csv("orders.csv")

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def dismiss_popup():
    """Closes page popup"""
    page = browser.page()
    page.click("button:text('OK')")

def fill_form_with_csv_data (orders):
    """Loads the orders data into the web"""
    for row in orders:
        dismiss_popup()
        fill_and_submit_orders(row)

def fill_and_submit_orders(order):
    """Fills in the orders data and click the "Order" buton"""
    page = browser.page()
    page.select_option("#head", order["Head"])
    page.click(f'#id-body-{order["Body"]}')
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])
    page.fill("#address", str(order["Address"]))
    page.click("#order")
    while page.is_visible("div.alert-danger"):
        page.click("#order")
    pdf_file = store_receipt_as_pdf(order["Order number"])
    screenshot = screenshot_robot(order["Order number"])
    embed_screenshot_to_receipt(screenshot, pdf_file)
    page.click("#order-another")

def store_receipt_as_pdf(order_number):
    """Print receipt and save it to pdf"""
    page = browser.page()
    order_receipt_html = page.locator("#order-completion").inner_html()
    pdf = PDF()
    file_name = f'output/receipts/order-receipt{order_number}.pdf'
    pdf.html_to_pdf(order_receipt_html, file_name)
    return file_name

def screenshot_robot(order_number):
    """Takes screenshot of the robot and saves it"""
    page = browser.page()
    file_name = f'output/robots/robot-screenshot{order_number}.png'
    robot_screenshot_html = page.locator("#robot-preview-image").screenshot(type='png', path=file_name)
    return file_name

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Add Image to PDF"""
    pdf = PDF()
    list_of_files = [
        pdf_file,
        screenshot,
    ]
    pdf.add_files_to_pdf(files=list_of_files,target_document=pdf_file)


def archive_receipts():
    """Zip recipts and archive them"""
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'output/archive.zip')
    