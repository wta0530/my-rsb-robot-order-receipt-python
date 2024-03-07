from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

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
        slowmo=100,
    )

    # embed_screenshot_to_receipt(10)


    open_robot_order_website()
    order_data = get_orders()
    for data in order_data:
        for order in data:
            close_the_annoying_modal()
            fill_the_form(order)
            preview_the_robot()
            submit_the_order()
            store_receipt_as_pdf(order["Order number"])
            screenshot_robot(order["Order number"])
            embed_screenshot_to_receipt(order["Order number"])
            go_to_order_another_robot()
    create_a_zip_file_of_the_receipts()


def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_the_annoying_modal():
    """Closing the annoying popup"""
    page = browser.page()
    page.wait_for_selector("//button[contains(text(),'OK')]")
    page.click("//button[contains(text(),'OK')]")
    page.wait_for_selector("//select[@id='head']")

def download_csv_test_data_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Downloads excel file from the given URL"""
    download_csv_test_data_file()
    csv_data = Tables()
    orders = csv_data.read_table_from_csv("orders.csv", header=True)
    order_data = csv_data.group_table_by_column(orders, "Order number")
    return order_data

def fill_the_form(order_data):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()

    page.wait_for_selector("#head")
    page.select_option("#head", order_data["Head"])
    page.check("//input[@name='body' and @value='"+order_data["Body"]+"']")
    page.fill("//input[@type='number']", order_data["Legs"])
    page.fill("#address", order_data["Address"])

def preview_the_robot():
    """Click the 'Preview' button"""
    page = browser.page()
    page.wait_for_selector("//button[@id='preview']")
    page.click("//button[@id='preview']")

def submit_the_order():
    """Click the 'Order' button and submit the order"""
    page = browser.page()
    page.wait_for_selector("//button[@id='order']")
    page.click("//button[@id='order']")

def fill_form_with_csv_data():
    """Read data from csv and fill in the sales form"""
    order_data = get_orders()
    for data in order_data:
        for order in data:
            fill_the_form(order)


def go_to_order_another_robot():
    """Click the 'Order' button and submit the order"""
    page = browser.page()
    # page.mouse.down()
    page.wait_for_selector("//button[@id='order-another']")
    page.click("//button[@id='order-another']")

def store_receipt_as_pdf(order_number):
    path = "output/"+order_number+".pdf"
    page = browser.page()
    page.wait_for_selector("//div[@id='receipt']")
    receipt_html = page.locator("//div[@id='receipt']").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(receipt_html, path)
    return path

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    path = "output/"+order_number+".png"
    page = browser.page()
    page.wait_for_selector("//div[@id='robot-preview-image']")
    page.locator("//div[@id='robot-preview-image']").screenshot(path=path)
    return path

def embed_screenshot_to_receipt(order_number):
    pdf = PDF()
    path = str(order_number)
    pdf.open_pdf("output/"+path+".pdf")
    pdf.add_watermark_image_to_pdf(image_path="output/"+path+".png", source_path="output/"+path+".pdf", output_path="output/"+path+".pdf")


def create_a_zip_file_of_the_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('./output', 'Receipts.zip', recursive=True, include = '*.pdf', exclude = '/.*')

    