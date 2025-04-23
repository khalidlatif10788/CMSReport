
import requests
from bs4 import BeautifulSoup
import pandas as pd

# Configuration
LOGIN_URL = "https://cms.must.edu.pk/sfms/checklogin.asp"
REPORT_URL = "https://cms.must.edu.pk/sfms/Reports/SemesterFeeCategory/Fee_Summary_Report.asp"
CREDENTIALS = {
    "username": "FEE",          # Replace with your username
    "password": "Eef@tomust"    # Replace with your password
}
SEMESTER_VALUE = "90-Spring 2025"  # Change this for different semesters
SCHEME_VALUE = "1"               # 1=Student Semester Fee

def main():
    # Start session
    with requests.Session() as session:
        # Step 1: Login
        print("⏳ Attempting login...")
        login_response = session.post(LOGIN_URL, data=CREDENTIALS)
        
        if "Login Failed" in login_response.text or login_response.url == LOGIN_URL:
            print("❌ Login failed. Check credentials or network connection.")
            return
        
        print("✅ Login successful")
        
        # Step 2: Load report page
        print("⏳ Loading report page...")
        report_response = session.get(REPORT_URL,timeout=60)
        report_soup = BeautifulSoup(report_response.text, 'html.parser')
        
        # Step 3: Prepare form data
        form_data = {
            "SelectSession": SEMESTER_VALUE,
            "SelectScheme": SCHEME_VALUE,
            "Submit": "Show Report"
        }
        
        # Step 4: Submit form with parameters
        print("⏳ Fetching fee data...")
        data_response = session.post(REPORT_URL, data=form_data)
        data_soup = BeautifulSoup(data_response.text, 'html.parser')
        
        # Step 5: Extract main data table
        data_table = data_soup.find('table', {'name': 'formtable'})
        
        if not data_table:
            print("❌ Data table not found. Possible issues:")
            print("- Semester/Scheme values need updating")
            print("- Page structure changed")
            with open('error_page.html', 'w', encoding='utf-8') as f:
                f.write(data_response.text)
            print("Saved error page to 'error_page.html'")
            return
        
        print("✅ Data table found")
        
        # Step 6: Process table data
        headers = [th.get_text(strip=True) for th in data_table.find('tr').find_all('td')]
        print(headers)
        rows = []
        science_department=["Biotechnology","Chemistry"]
        engineering_department=["Civil Engineering","Computer Systems Engineering"]
        
        faculties={}
        
        faculty_sceneces_departments={}
        faculty_Engineering_departments={}

        for row in data_table.find_all('tr')[1:]:  # Skip header row
            
            
            # print(f" Department Name: {row.find_all('td')[0]} Fee Collected: {row.find_all('td')[4]}")
            if row.find_all('td')[0].get_text(strip=True) in science_department:
                faculty_sceneces_departments[row.find_all('td')[0].get_text(strip=True)]={headers[4]:row.find_all('td')[4].get_text(strip=True)}
                faculties["Sciences"]=faculty_sceneces_departments
            elif row.find_all('td')[0].get_text(strip=True) in engineering_department:
                faculty_Engineering_departments[row.find_all('td')[0].get_text(strip=True)]={headers[4]:row.find_all('td')[4].get_text(strip=True)}
                faculties["Engineering"]=faculty_Engineering_departments
        print(faculties)
            
       

if __name__ == "__main__":
    main()
