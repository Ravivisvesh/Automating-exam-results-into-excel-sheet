from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import csv

# Chrome options
options = Options()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--allow-insecure-localhost')

# Use webdriver-manager to auto-manage ChromeDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 10)

# Read input CSV
with open('Book1test.csv', 'r', encoding='utf-8-sig') as boo:
    reader = csv.DictReader(boo)
    rows = list(reader)

output_data = []

for i in rows:
    driver.get("https://mcetresultpage.netlify.app/")
    driver.find_element(By.CSS_SELECTOR, "input[value='STUDENTS']").click()

    roll = i["roll"].strip()
    dob = i["birth"].strip()

    # Wait for elements
    cont = wait.until(EC.presence_of_element_located((By.ID, "txtLoginId")))
    pas1 = wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))
    
    cont.send_keys(roll)
    pas1.send_keys(dob)

    # Captcha handling
    cap_display = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "noselect")))
    captcha_text = cap_display.text.strip()
    captcha_input = wait.until(EC.presence_of_element_located((By.ID, "txtCaptcha")))
    captcha_input.send_keys(captcha_text)

    # Submit login
    driver.find_element(By.XPATH, "//input[@type='button' and @value='Submit']").click()

    # Navigate to results
    result_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[value='End Sem Results']")))
    result_button.click()

    # Switch to results frame
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "divRightPane")))
    rows_table = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[@class='tblBRDefault']//tr")))

    for row in rows_table[1:]:
        tds = row.find_elements(By.TAG_NAME, "td")
        if len(tds) >= 8:
            coursecode = tds[2].text.strip()
            gr = tds[6].text.strip()
            output_data.append({"rollnumber": roll, "coursecode": coursecode, "gr": gr})

    driver.switch_to.default_content()

driver.quit()

# Write output CSV
with open('results_output.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=["rollnumber", "coursecode", "gr"])
    writer.writeheader()
    writer.writerows(output_data)
