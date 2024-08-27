import re, os, io, time, docx, shutil, win32com.client, pythoncom
from tqdm.auto import tqdm
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from datetime import datetime
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

def get_user_input():
    email_dict = {
        "e": "e.desinety@sycomore-am.com",
        "t": "tony.lebon@sycomore-am.com",
        "a": "alex.mory@sycomore-am.com"
    }
    company_name = input("Enter the company name as saved in the SSC folder: ")
    deal_entry_code = input("Enter the Deal Roadshow Entry Code: ")
    
    print("Email options:")
    for key, value in email_dict.items():
        print(f"Enter '{key}' for {value}")
    
    email_short = input("DealRoadshow email address: ").lower()
    email = email_dict.get(email_short, email_short)

    mode = int(input("Enter 1 for 'debug mode' (keep Word doc and screenshots) or 2 for 'user mode' (only pdf file): "))
    return company_name, deal_entry_code, email, mode

def initialize_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("start-maximized")
    options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    print("Initializing the WebDriver...")
    return driver

def login_to_dealroadshow(driver, email, deal_entry_code):
    try:
        print("Opening the Deal Roadshow login page...")
        login_url = 'https://dealroadshow.finsight.com/'
        driver.get(login_url)
        time.sleep(1)
        driver.find_element(By.XPATH, "//input[@name='email']").send_keys(email)
        driver.find_element(By.XPATH, "//input[@name='entryCode']").send_keys(deal_entry_code, Keys.RETURN)
        time.sleep(7)
        
        if (driver.current_url == login_url+'e/'+deal_entry_code):
            print("Login successful.")
            return True
        else:
            print("Login failed.")
            return False
            
    except (NoSuchElementException, TimeoutException) as e:
        print(f"Error: {e}")
        return False

def wait_for_interaction(driver, max_attempts=12, delay=5):
    attempts = 0
    while attempts < max_attempts:
        try:
            agree_button = driver.find_element(By.XPATH, "//button[span[text()='I Agree']]")
            agree_button.click()
            time.sleep(5)
            return True
        except NoSuchElementException:
            print("No 'I Agree' button found. Please interact with any modals. Retrying in 5 seconds...")
            time.sleep(delay)
            attempts += 1
    print("Max attempts reached.")
    return False

def pause_video(driver):
    try:
        wait = WebDriverWait(driver, 10)
        video_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "video.vjs-tech[playsinline][role='application']"))
                                  )
        driver.execute_script("arguments[0].pause();", video_element)
        time.sleep(3)
    except NoSuchElementException as e:
        print("No Video element found")
    except TimeoutException as e:
        print(f"Error: Timeout while trying to find element. Details: {e}")

def navigate_to_first_slide(driver):
    print("Navigating to the first slide...")
    driver.find_element(By.XPATH, "//button[@data-test='toggleMenuButton']").click()
    while True:
        previous_slide_button = driver.find_element(By.XPATH, "//button[@data-test='viewerPreviousSlideButton']")
        if previous_slide_button.get_attribute("disabled") == 'true':
            break
        previous_slide_button.click()
        time.sleep(1)

def get_nb_slides(driver):
    try:
        return int(driver.find_element(By.XPATH, "//div[contains(@class, 'actionBar_totalSlidesNum__')]").text.strip())
    except (NoSuchElementException, ValueError):
        print("Error getting total slides.")
        return None

def create_output_folders(company_name):
    print("Creating output folders for saving screenshots...")
    current_time = datetime.now().strftime('%Y-%m-%d')
    output_folder = os.path.join("F:\\GESTION\\SSC", company_name, f'{current_time} NRS {company_name}')
    screenshots_folder = os.path.join(output_folder, f'screenshots_{company_name}')
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(screenshots_folder, exist_ok=True)
    return screenshots_folder, output_folder

def take_screenshots(driver, screenshots_folder, nb_slides, progress_callback=None):
    print("Starting roadshow scraping...")
    image_index = 1
    screenshot_paths = []
    driver.set_window_size(1920,1080)
    time.sleep(3)

    progress_bar = None
    if nb_slides is not None:
        progress_bar = tqdm(total=nb_slides, desc='Taking screenshots', unit='slide')
    while True:
        try:
            screenshot_path = os.path.join(screenshots_folder, f'img_{image_index}.png')
            driver.find_element(By.XPATH, "//div[@data-test='slidesViewerSlide']").screenshot(screenshot_path)
            screenshot_paths.append(screenshot_path)
            image_index += 1
            if progress_bar:
                progress_bar.update(1)
            if progress_callback:
                    progress_callback(image_index - 1)
                
            next_slide_button = driver.find_element(By.XPATH, "//button[@data-test='viewerNextSlideButton']")
            if next_slide_button.get_attribute("disabled") == 'true':
                print(f"Scraping completed. {image_index - 1} screenshots saved in {screenshots_folder}")
                break
            next_slide_button.click()
            time.sleep(2)
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Error: {e}")
            break
    return screenshot_paths

def save_screenshots_to_word(screenshot_paths, output_folder, company_name):
    print("Starting screenshot compilation in a Word Docx...")
    current_time = datetime.now().strftime('%Y-%m-%d')
    doc = Document()

    with Image.open(screenshot_paths[0]) as img:
        img_width, img_height = img.size
    aspect_ratio = img_width / img_height
    
    # Set the document orientation to landscape
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width = section.page_width
    new_height = new_width / aspect_ratio
    section.page_height = int(new_height)
    
    # Set margins to zero
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Pt(0)
    
    for screenshot_path in screenshot_paths:
        doc.add_picture(screenshot_path, width=new_width, height=new_height)
    doc_path = os.path.join(output_folder, f'{current_time} NRS {company_name}.docx')
    doc.save(doc_path)
    print(f'Word document saved as {doc_path}')
    return doc_path

def save_screenshots_to_pdf(doc_path):
    print("Saving docx as pdf...")
    pdf_path = os.path.splitext(doc_path)[0] + '.pdf'
    try:
        word = win32com.client.gencache.EnsureDispatch('Word.Application', pythoncom.CoInitialize())
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17) #wdFormatPDF
        doc.Close()
        word.Quit()
        print(f'PDF document saved as {pdf_path}')
    except Exception as e:
        print("An error occurred:", e)
    finally:    
        word = None
    return pdf_path

def cleanup_files(output_folder, pdf_path, mode):
    if mode == 2:
        print("Cleaning NRS folder...")
        parent_folder = os.path.dirname(output_folder)
        new_pdf_path = os.path.join(parent_folder, os.path.basename(pdf_path))
        shutil.move(pdf_path, new_pdf_path)
        print(f'Moved PDF to {new_pdf_path}')
        
        if os.path.exists(output_folder):
            shutil.rmtree(output_folder)
            print(f'Deleted {output_folder}')