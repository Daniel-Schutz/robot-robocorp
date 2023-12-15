from RPA.FileSystem import FileSystem
from robocorp.tasks import task
from robocorp import browser
from RPA.PDF import PDF
from RPA.HTTP import HTTP
from RPA.Tables import Tables
import time
from RPA.PDF import PDF
import zipfile
import os


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
        slowmo=700,
    )
    open_robot_order_website()
    page = browser.page()
    page.locator("div#order-completion")
    get_orders()

    read_csv()

    archive_receipts()


def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    close_annoying_modal()


def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(
        url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def read_csv():
    library = Tables()
    data = library.read_table_from_csv("orders.csv", header=True)
    for row in data:
        fill_the_form(row)


def close_annoying_modal():
    page = browser.page()
    page.click("text=OK")


def fill_the_form(row):
    """Fills in the form and click the 'Submit' button"""
    page = browser.page()

    page.select_option("#head", str(row["Head"]))
    page.click(f"#id-body-{str(row['Body'])}")
    page.fill('input[type="number"]', str(row["Legs"]))
    page.fill("#address", str(row["Address"]))
    page.click("text=Preview")
    time.sleep(1)
    order()
    page.click("button#order-another.btn.btn-primary")
    close_annoying_modal()


def order():
    page = browser.page()
    page.click("button#order.btn.btn-primary")
    time.sleep(2)
    while (page.locator("div#order-completion").count() == 0):
        page.click("button#order.btn.btn-primary")
     # Loop until the order is submitted successfully or the retry limit is reached
    order_number = page.locator("p.badge.badge-success").inner_text()
    print("Numeroooo", order_number)
    pdf_file = store_receipt_as_pdf(order_number)
    screenshot = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot, pdf_file)


def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    receipt_html = page.locator("div#receipt.alert.alert-success").inner_html()
    pdf = PDF()
    path = f"output/{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, path)

    return path


def screenshot_robot(order_number):
    page = browser.page()
    path = f"output/{order_number}.png"
    page.screenshot(path=path)
    return path


def embed_screenshot_to_receipt(screenshot, pdf_file):
    # Use the RPA.PDF library to append the screenshot to the PDF
    pdf = PDF()
    pdf.add_files_to_pdf(files=[pdf_file, screenshot],
                         target_document=pdf_file, append=True)


def archive_receipts():
    folder_path = r'C:\Users\danie\Doxci2\Robocorp\course2\output'
    zip_file_name = os.path.join(
        folder_path, 'receipts.zip')  # Save inside the folder
    with zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(
                        file_path, folder_path))
