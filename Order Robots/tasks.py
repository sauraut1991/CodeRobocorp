from robocorp.tasks import task
from robocorp import browser 
from RPA.HTTP import HTTP
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
        slowmo=500,
    )
    open_robot_order_website()
    download_robot_sales_data()
    fill_form_with_csv_data()
    archive_receipts()
  
def open_robot_order_website():
    
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_robot_sales_data():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def get_robot_orders_from_excel():
    Table = Tables()
    worksheet = Table.read_table_from_csv("orders.csv")

    return worksheet


def fill_form_with_csv_data():

    page=browser.page()

    orders = get_robot_orders_from_excel()

    for row in orders:
        page.click("text=OK")
        page.select_option("#head",row["Head"])
        page.check("#id-body-"+str(row["Head"]))
        page.fill("//input[@class='form-control']",str(row["Legs"]))
        page.fill('//input[contains(@id, "address")]',str(row["Address"]))
        
        page.click("#order")

        if(page.is_visible('//div[@class="alert alert-danger"]')):
             page.click("#order")

        store_receipt_as_pdf(row["Order number"])
        screenshot_robot(row["Order number"])  
        page.click("#order-another")   

def store_receipt_as_pdf(ordernumber):
    page=browser.page()
    order_receipt_html= page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(order_receipt_html, "output/sales_results-OrderNumber - "+ordernumber+".pdf")

def screenshot_robot(order_number):
          page = browser.page()
          screenshot_path="output/Robot-Image Order Number - "+order_number+".png"
          page.screenshot(path="output/Robot-Image Order Number - "+order_number+".png")

          pdf_file_path = "output/sales_results-OrderNumber - "+order_number+".pdf"

          embed_screenshot_to_receipt(screenshot_path,pdf_file_path)

def embed_screenshot_to_receipt(screenshot_path,pdf_file_path):
     pdf = PDF()
     list_of_files = [
        pdf_file_path,
        screenshot_path
    ]
     pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=(pdf_file_path))
     
def archive_receipts():
     archive = Archive()
     archive.archive_folder_with_zip("./output","RobotOrderDetails.zip",False,include='*.pdf')

     


    




           


