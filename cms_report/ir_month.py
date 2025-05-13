from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time

LOGIN_URL = "https://cms.must.edu.pk/sfms/checklogin.asp"
REPORT_URL = "https://cms.must.edu.pk/sfms/Reports/SemesterFeeCategory/SelectCriteriaAllbank.asp?wlink=FeeReconciliationReport_for_Selected_Classallbank.asp"
CREDENTIALS = {
    "login": "FEE",
    "password": "Eef@tomust"
}

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(10)
    return driver

def login(driver):
    print("Attempting login...")
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 10)

    try:
        username_field = wait.until(EC.presence_of_element_located((By.NAME, "login")))
        username_field.send_keys(CREDENTIALS["login"])

        password_field = wait.until(EC.presence_of_element_located((By.NAME, "password")))
        password_field.send_keys(CREDENTIALS["password"])
        password_field.send_keys(Keys.RETURN)

        wait.until(EC.url_changes(LOGIN_URL))
        if driver.current_url == LOGIN_URL:
            raise Exception("Login failed. Check credentials or network connection.")
        print("Login successful")

    except Exception as e:
        print(f"Login failed due to: {e}")
        driver.quit()
        raise

def wait_for_full_table_load(driver, timeout=120, stable_check_delay=5):
    print("Waiting for full table to load...")
    previous_count = 0
    start_time = time.time()
    same_count_repeat = 0

    while True:
        try:
            rows = driver.find_elements(By.XPATH, "//table[@id='AutoNumber2']//tr")
            current_count = len(rows)

            print(f"Found {current_count} rows so far...")

            if current_count == previous_count:
                same_count_repeat += 1
            else:
                same_count_repeat = 0  # reset if count changes

            if same_count_repeat >= 3 and current_count > 10:  # considered stable
                print(f"Table stabilized with {current_count} rows.")
                break

            # Scroll to bottom in case of lazy loading
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            previous_count = current_count
            time.sleep(stable_check_delay)

            if time.time() - start_time > timeout:
                print(f"Timeout reached with {current_count} rows.")
                break

        except Exception as e:
            print(f"Error while waiting for table: {e}")
            break


def fetch_report(driver):
    print("Loading report page...")
    driver.get(REPORT_URL)
    wait = WebDriverWait(driver, 500)  # Increased timeout

    try:
        # Wait for form to load
        wait.until(EC.presence_of_element_located((By.NAME, "SelectClass")))

        # Select dropdown values
        Select(driver.find_element(By.NAME, "SelectClass")).select_by_value("ALL")
        Select(driver.find_element(By.NAME, "SelectSession")).select_by_value("-1")
        Select(driver.find_element(By.NAME, "AcademicStatus")).select_by_value("=1")
        Select(driver.find_element(By.NAME, "SelectScheme")).select_by_value("1")
        Select(driver.find_element(By.NAME, "selectBank")).select_by_value("-1")

        # Handle dates carefully
        # Start date
        start_date_field = wait.until(EC.element_to_be_clickable((By.NAME, "StartDate")))
        driver.execute_script("arguments[0].removeAttribute('readonly')", start_date_field)
        start_date_field.clear()
        start_date_field.send_keys("1/1/2025")

        # End date
        end_date_field = wait.until(EC.element_to_be_clickable((By.NAME, "EndDate")))
        driver.execute_script("arguments[0].removeAttribute('readonly')", end_date_field)
        end_date_field.clear()
        end_date_field.send_keys("31/1/2025")


        # Verify dates were entered correctly
        print(f"Start Date: {start_date_field.get_attribute('value')}")
        print(f"End Date: {end_date_field.get_attribute('value')}")

        # Submit form using JavaScript click to avoid interception
        submit_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@value='Continue']")))
        driver.execute_script("arguments[0].click();", submit_button)

        # Wait for report to load
        wait.until(EC.presence_of_element_located((By.ID, "AutoNumber2")))
        wait_for_full_table_load(driver)

        return driver.page_source

    except Exception as e:
        print(f"Error in fetch_report: {str(e)}")
        driver.save_screenshot("report_error.png")
        raise
def process_data(html):
    soup = BeautifulSoup(html, 'html.parser')
    data_table = soup.find('table', {'id': 'AutoNumber2'})

    if not data_table:
        print("Data table not found")
        return None
    headers = [(index, th.get_text(strip=True)) 
           for index, th in enumerate(data_table.find('tr').find_all('td'), start=1)]
    print(headers)
    all_departments = {
        "Secretariat": ["Secretariat"], "Examinations": ["Examinations"], "DSA": ["DSA"], "Hostels": ["Hostels"],
        "Akson": ["Akson"], "Affilation": ["Affilation"], "Electrical": ["BEE"], "Mechanica": ["BME"], "CSE": ["CSE"],
        "Civil": ["BCV"], "Software": ["BSE"], "ESE": [""], "ISE": [""], "MIT": ["BCT","BAI","BMT","IET","BAT"], "CS&IT": ["BIT","BCS"], "Physics": [""],
        "Chemistry": ["RCH","BCH"], "Mathematics": ["BSM"], "Botany": [""], "Zoology": ["BZO"], "Bio-Tech": ["BBT","RBT"], "Statistics": [""],
        "Home Economics": [""], "Fashion & Design": [""], "IIS": ["BIS"], "Arabic": [""], "English": ["ENG"],
        "Economics": [""], "Law": ["LLB"], "Education": ["BED"], "International Relation": [""], "Sociology": [""],
        "Media Com": [""], "BA": [""], "Comm": ["BCM"], "B&F": [""], "Accounting & Finance": ["BFA"], "HRM": [""],
        "Forestry (PC)": ["BFT"], "Tourism (PC)": ["BTH"], "Education (PC)": ["EDE"], "International Relation (PC)": ["BIR"],
        "Physics (PC)": ["BPH"], "Pharmacy": ["DPH"], "Human Nutrition & Dietetics": [""], "Doctor of Physical Therapy": ["DPT"],
        "Medical Lab Technology": ["MLT","OTT"], "Microbiology": [""], "DVM": [""], "BS Animal Sciences": [""],
        "Livestock Assistant Diploma": [""], "Pharmacy": [""],
    }

    fee_summary = {
    dept: {
        "Registration Fee":0,
        "Admission Fee (at the time of Admission)":0,
        "Library Security":0,
        "Laboratory Security":0,
        "Endowment Fund":0,
        "Student Welfare Fund":0,
        "Tuition Fee (Subsidized Students)":0,
        "Tuition Fee (Regular Students)":0,
        "Sports Fee":0,
        "Library Fee":0,
        "Laboratory Fee":0,
        "Course Registration Fee":0,
        "Health Care Fee":0,
        "Research Fund":0,
        "Fuel / Generator Charges":0,
        "Bus Charges":0,
        "Examination Fee":0,
        "IT Fund":0,
       
        "Internet Charges":0,
        "Student Insurance Fee":0,
        "Security Charges":0,
        "Other Charges/ Field Work":0,
        "Thesis/ Project Fee":0,
        "Degree Fee (Enrolled Students)":0,
        "Student Card Fee per semester":0,
        "Transport Development Fund":0,
        "Campus Management System Support Fund":0,
        "Survey Camp Fee":0,
        "Development Fund":0,
        "Hospital Visit Fee":0,
        "Legal Fee/LAW Charges":0,
        "Fines (Late Fee Fine etc.)":0
    }
    for dept in all_departments
}
    fee_summary["Other"] = {
     "Registration Fee":0,
        "Admission Fee (at the time of Admission)":0,
        "Library Security":0,
        "Laboratory Security":0,
        "Endowment Fund":0,
        "Student Welfare Fund":0,
        "Tuition Fee (Subsidized Students)":0,
        "Tuition Fee (Regular Students)":0,
        "Sports Fee":0,
        "Library Fee":0,
        "Laboratory Fee":0,
        "Course Registration Fee":0,
        "Health Care Fee":0,
        "Research Fund":0,
        "Fuel / Generator Charges":0,
        "Bus Charges":0,
        "Examination Fee":0,
        "IT Fund":0,
        "Masque Fund ":0,
        "Internet Charges ":0,
        "Student Insurance Fee ":0,
        "Security Charges":0,
        "Other Charges/ Field Work":0,
        "Thesis/ Project Fee":0,
        "Degree Fee (Enrolled Students)":0,
        "Student Card Fee per semester":0,
        "Transport Development Fund":0,
        "Campus Management System Support Fund":0,
        "Survey Camp Fee":0,
        "Development Fund":0,
        "Hospital Visit Fee":0,
        "Legal Fee/LAW Charges":0,
        "Fines (Late Fee Fine etc.)":0
}

    row_count = 0
    processed_rows = 0

    for row in data_table.find_all('tr')[2:]:
        row_count += 1
        cells = row.find_all('td')

        if len(cells) < 17:
            continue

        try:
            current_program = cells[3].get_text(strip=True)
            if not current_program:
                continue

            matched_dept = "Other"
            for dept, programs in all_departments.items():
                if current_program in programs or not programs:
                    matched_dept = dept
                    break

            registration_str = cells[67].get_text(strip=True).replace(',', '')
            registration = int(registration_str) if registration_str.isdigit() else 0
            # security
            admission_str = cells[66].get_text(strip=True).replace(',', '')
            admission= int() if admission_str.isdigit() else 0
            # security
            library_security_str = cells[44].get_text(strip=True).replace(',', '')
            library_security = int(library_security_str) if library_security_str.isdigit() else 0
            # security
            Laboratory_security_str = cells[39].get_text(strip=True).replace(',', '')
            Laboratory_security = int(Laboratory_security_str) if Laboratory_security_str.isdigit() else 0
            # security
            Endowment_Fund_str = cells[25].get_text(strip=True).replace(',', '')
            Endowment_Fund = int(Endowment_Fund_str) if Endowment_Fund_str.isdigit() else 0
            # security
            Student_Welfare_Fund_str = cells[79].get_text(strip=True).replace(',', '')
            Student_Welfare_Fund = int(Student_Welfare_Fund_str) if Student_Welfare_Fund_str.isdigit() else 0
             # security
            Tuition_Fee_Subsidized_Students_str = cells[89].get_text(strip=True).replace(',', '')
            Tuition_Fee_Subsidized_Students = int(Tuition_Fee_Subsidized_Students_str) if Tuition_Fee_Subsidized_Students_str.isdigit() else 0
            # security
            # security
            Non_Subsidized_Regular_students_str = cells[49].get_text(strip=True).replace(',', '')
            Non_Subsidized_Regular_students = int(Non_Subsidized_Regular_students_str) if Non_Subsidized_Regular_students_str.isdigit() else 0
           
            Sports_Fee_str = cells[74].get_text(strip=True).replace(',', '')
            Sports_Fee = int(Sports_Fee_str) if Sports_Fee_str.isdigit() else 0
            # security
            Library_Fee_str = cells[43].get_text(strip=True).replace(',', '')
            Library_Fee = int(Library_Fee_str) if Library_Fee_str.isdigit() else 0
             # security
            Laboratory_Fee_str = cells[38].get_text(strip=True).replace(',', '')
            Laboratory_Fee = int(Laboratory_Fee_str) if Laboratory_Fee_str.isdigit() else 0
            # security
            Course_Registration_fee_str= cells[21].get_text(strip=True).replace(',', '')
            Course_Registration_fee = int(Course_Registration_fee_str) if Course_Registration_fee_str.isdigit() else 0
            # security
            Health_Care_Fee_str = cells[34].get_text(strip=True).replace(',', '')
            Health_Care_Fee = int(Health_Care_Fee_str) if Health_Care_Fee_str.isdigit() else 0
            # security
            Research_Fund_str = cells[69].get_text(strip=True).replace(',', '')
            Research_Fund = int(Research_Fund_str) if Research_Fund_str.isdigit() else 0
            # security
            Fuel_Generator_Charges_str = cells[32].get_text(strip=True).replace(',', '')
            Fuel_Generator_Charges = int(Fuel_Generator_Charges_str) if Fuel_Generator_Charges_str.isdigit() else 0
            # security
            Bus_Charges_str = cells[16].get_text(strip=True).replace(',', '')
            Bus_Charges = int(Bus_Charges_str) if Bus_Charges_str.isdigit() else 0
            # security
            Examination_Fee_str = cells[26].get_text(strip=True).replace(',', '')
            Examination_Fee = int(Examination_Fee_str) if Examination_Fee_str.isdigit() else 0
            # security
            IT_Fund_str = cells[35].get_text(strip=True).replace(',', '')
            IT_Fund = int(IT_Fund_str) if IT_Fund_str.isdigit() else 0
            # security
            Mosque_Fund_str = cells[48].get_text(strip=True).replace(',', '')
            Mosque_Fund = int(Mosque_Fund_str) if Mosque_Fund_str.isdigit() else 0
             # security
            Internet_Charges_str = cells[36].get_text(strip=True).replace(',', '')
            Internet_Charges = int(Internet_Charges_str) if Internet_Charges_str.isdigit() else 0
            # security
            Student_Insurance_Fee_str = cells[77].get_text(strip=True).replace(',', '')
            Student_Insurance_Fee = int(Student_Insurance_Fee_str) if Student_Insurance_Fee_str.isdigit() else 0
            # security
            Security_Charges_str = cells[71].get_text(strip=True).replace(',', '')
            Security_Charges = int(Security_Charges_str) if Security_Charges_str.isdigit() else 0
            # security
            Other_Charges_Field_Work_str = cells[29].get_text(strip=True).replace(',', '')
            Other_Charges_Field_Work = int(Other_Charges_Field_Work_str) if Other_Charges_Field_Work_str.isdigit() else 0
            # security
            Thesis_Project_Fee_str = cells[82].get_text(strip=True).replace(',', '')
            Thesis_Project_Fee = int(Thesis_Project_Fee_str) if Thesis_Project_Fee_str.isdigit() else 0
            # security
            Degree_Fee_str = cells[23].get_text(strip=True).replace(',', '')
            Degree_Fee = int(Degree_Fee_str) if Degree_Fee_str.isdigit() else 0
            # security
            Student_Card_Fee_str = cells[76].get_text(strip=True).replace(',', '')
            Student_Card_Fee = int(Student_Card_Fee_str) if Student_Card_Fee_str.isdigit() else 0
            # security
            Transport_Development_Fund_str = cells[86].get_text(strip=True).replace(',', '')
            Transport_Development_Fund = int(Transport_Development_Fund_str) if Transport_Development_Fund_str.isdigit() else 0
            # security
            CMS_Support_Fund_str = cells[18].get_text(strip=True).replace(',', '')
            CMS_Support_Fund = int(CMS_Support_Fund_str) if CMS_Support_Fund_str.isdigit() else 0
            # security
            Survey_Camp_Fee_str = cells[81].get_text(strip=True).replace(',', '')
            Survey_Camp_Fee = int(Survey_Camp_Fee_str) if Survey_Camp_Fee_str.isdigit() else 0
            # security
            fine_str = cells[30].get_text(strip=True).replace(',', '')
            fine = int(fine_str) if fine_str.isdigit() else 0
           

            fee_summary[matched_dept]["Registration Fee"] += registration
            fee_summary[matched_dept]["Admission Fee (at the time of Admission)"] += admission
            fee_summary[matched_dept]["Library Security"] += library_security
            fee_summary[matched_dept]["Laboratory Security"] += Laboratory_Fee
            fee_summary[matched_dept]["Endowment Fund"] += Endowment_Fund
            fee_summary[matched_dept]["Student Welfare Fund"] += Student_Welfare_Fund
            fee_summary[matched_dept]["Tuition Fee (Subsidized Students)"] += Tuition_Fee_Subsidized_Students
            fee_summary[matched_dept]["Tuition Fee (Regular Students)"] += Non_Subsidized_Regular_students
            fee_summary[matched_dept]["Sports Fee"] += Sports_Fee
            fee_summary[matched_dept]["Library Fee"] += Library_Fee
            fee_summary[matched_dept]["Laboratory Fee"] += Laboratory_Fee
            fee_summary[matched_dept]["Course Registration Fee"] += Course_Registration_fee
            fee_summary[matched_dept]["Health Care Fee"] += Health_Care_Fee
            fee_summary[matched_dept]["Research Fund"] += Research_Fund
            fee_summary[matched_dept]["Fuel / Generator Charges":] += Fuel_Generator_Charges
            fee_summary[matched_dept]["Bus Charges"] += Bus_Charges
            fee_summary[matched_dept]["Examination Fee"] += Examination_Fee
            fee_summary[matched_dept]["IT Fund"] += IT_Fund
            # fee_summary[matched_dept]["Masque Fund"] += Mosque_Fund
            fee_summary[matched_dept]["Internet Charges"] += Internet_Charges
            fee_summary[matched_dept]["Student Insurance Fee"] += Student_Insurance_Fee
            fee_summary[matched_dept]["Security Charges"] += Security_Charges
            fee_summary[matched_dept]["Other Charges/ Field Work"] += Other_Charges_Field_Work
            fee_summary[matched_dept]["Thesis/ Project Fee"] += Thesis_Project_Fee
            fee_summary[matched_dept]["Degree Fee (Enrolled Students)"] += Degree_Fee
            fee_summary[matched_dept]["Student Card Fee per semester"] += Student_Card_Fee
            fee_summary[matched_dept]["Transport Development Fund"] += Transport_Development_Fund
            fee_summary[matched_dept]["Campus Management System Support Fund"] += CMS_Support_Fund
            fee_summary[matched_dept]["Survey Camp Fee"] += Survey_Camp_Fee
            # fee_summary[matched_dept]["Development Fund"] += fee
            # fee_summary[matched_dept]["Hospital Visit Fee"] += fee
            # fee_summary[matched_dept]["Legal Fee/LAW Charges"] += fee
            fee_summary[matched_dept]["Fines (Late Fee Fine etc.)"] += fine
            
            processed_rows += 1

        except Exception as e:
            continue

    print(f"\nTotal rows scanned: {row_count}")
    print(f"Rows with valid data: {processed_rows}")

    # print("\nDepartment-wise Bus Charges:")
    # for dept in sorted(fee_summary):
    #     print(f"\nDepartment: {dept}")
    #     print("-------------------------")
    #     for fee_name, amount in fee_summary[dept].items():
    #         print(f"{fee_name}: {amount}")
    #     print("-------------------------")
   
    import csv

    # Assuming fee_summary is your dictionary
    all_departments = list(fee_summary.keys())
    fee_heads = list(fee_summary[all_departments[0]].keys()) if all_departments else []

    with open('income_report.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        
        # Write header row
        writer.writerow(['Fee Head'] + all_departments)
        
        # Write each fee head row
        for fee_head in fee_heads:
            row = [fee_head]
            for dept in all_departments:
                row.append(fee_summary[dept][fee_head])
            writer.writerow(row)
def main():
    driver = setup_driver()
    try:
        login(driver)
        html = fetch_report(driver)
        process_data(html)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()

       # Initialize counters and debug info
        #     elif row.find_all('td')[0].get_text(strip=True) in engineering_departments:
        #         faculty_Engineering[row.find_all('td')[0].get_text(strip=True)]={headers[4]:row.find_all('td')[4].get_text(strip=True),headers[11]:row.find_all('td')[11].get_text(strip=True),headers[12]:row.find_all('td')[12].get_text(strip=True)}
        #         faculties["Faculty of Engineering and Technology"]=faculty_Engineering
        #     elif row.find_all('td')[0].get_text(strip=True) in social_sciences_departments:
        #         faculty_social_sciences[row.find_all('td')[0].get_text(strip=True)]={headers[4]:row.find_all('td')[4].get_text(strip=True),headers[11]:row.find_all('td')[11].get_text(strip=True),headers[12]:row.find_all('td')[12].get_text(strip=True)}
        #         faculties["Faculty of Social Sciences and Humanities"]=faculty_social_sciences
        #     elif row.find_all('td')[0].get_text(strip=True) in health_medical_departments:
        #         faculty_health_medical[row.find_all('td')[0].get_text(strip=True)]={headers[4]:row.find_all('td')[4].get_text(strip=True),headers[11]:row.find_all('td')[11].get_text(strip=True),headers[12]:row.find_all('td')[12].get_text(strip=True)}
        #         faculties["Faculty of Health and Medical Sciences"]=faculty_health_medical
        #     elif row.find_all('td')[0].get_text(strip=True) in must_businessSchool_departments:
        #         faculty_must_business[row.find_all('td')[0].get_text(strip=True)]={headers[4]:row.find_all('td')[4].get_text(strip=True),headers[11]:row.find_all('td')[11].get_text(strip=True),headers[12]:row.find_all('td')[12].get_text(strip=True)}
        #         faculties["Faculty of MBS"]=faculty_must_business
        
        # create_outstanding_dues_report(faculties, "outstanding_dues_report_Spring25.csv")
        # print("CSV file created successfully! Open it in Excel.")
            
# def create_outstanding_dues_report(data_dict, output_filename):
    
#     csv_lines = []
    
#     # Add title
#     csv_lines.append(',"Outstanding Dues Spring Semester 2025 As Per Student Fee Management Record",,,,,')
    
#     # Add headers
#     csv_lines.append('Sr. No.,Faculty,Department,Collected Dues,Outstanding Dues,No. Of Student Outstanding Dues')
    
   
#     sr_no = 1
#     for faculty, departments in data_dict.items():
#         first_dept = True
        
#         for dept, values in departments.items():
#             faculty_name = faculty if first_dept else ""
            
            
#             collected = values.get('Fee Collected', '0').replace(',', '')
#             outstanding = values.get('Outstanding Dues', '0').replace(',', '')
#             students = values.get('No. of Students, Outstanding Dues', '0').replace(',', '')
            
#             row = [
#                 str(sr_no),
#                 f'"{faculty_name}"',
#                 f'"{dept}"',
#                 collected,
#                 outstanding,
#                 students
#             ]
#             csv_lines.append(','.join(row))
#             sr_no += 1
#             first_dept = False
    
    
#     last_data_row = sr_no + 1  # +1 for header row
#     csv_lines.append(f'"Total outstanding Amount",,,,=SUM(E3:E{last_data_row}),=SUM(F3:F{last_data_row})')
    
    
#     with open(output_filename, 'w', encoding='utf-8') as f:
#         f.write('\n'.join(csv_lines))




# if __name__ == "__main__":
#     main()
