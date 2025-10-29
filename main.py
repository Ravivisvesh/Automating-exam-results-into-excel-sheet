from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import csv
import time

# Chrome options
options = Options()
options.add_argument("--ignore-certificate-errors")
options.add_argument("--allow-insecure-localhost")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)

# Read input CSV
with open('Book1test.csv', 'r', encoding='utf-8-sig') as boo:
    reader = csv.DictReader(boo)
    rows = list(reader)

output_data = []

for i in rows:
    driver.get("https://mcetresultpage.netlify.app/")
    roll = i["roll"].strip()
    dob = i["birth"].strip()
    print(f"Fetching results for Roll: {roll}, DOB: {dob}")

    # Wait for login fields
    wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(roll)
    wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(dob)

    # Click login
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "button3"))).click()

    # Wait for the result table to appear
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "card")))

    # Give a small buffer for React rendering
    time.sleep(1)

    # Get all rows of results
    rows_table = driver.find_elements(By.CSS_SELECTOR, "tr.item")

    for row in rows_table:
        tds = row.find_elements(By.TAG_NAME, "td")
        if len(tds) >= 4:
            subject = tds[1].text.strip()
            grade = tds[2].text.strip()
            grade_points = tds[3].text.strip()
            output_data.append({
                "rollnumber": roll,
                "subject": subject,
                "grade": grade,
                "grade_points": grade_points
            })

driver.quit()

# Write output CSV
with open('results_output.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=["rollnumber", "subject", "grade", "grade_points"])
    writer.writeheader()
    writer.writerows(output_data)

