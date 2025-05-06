
import requests
from bs4 import BeautifulSoup
import pandas as pd
#Convert Excel work
# from openpyxl import Workbook
# from openpyxl.styles import Font, Alignment


LOGIN_URL = "https://cms.must.edu.pk/sfms/checklogin.asp"
REPORT_URL ="https://cms.must.edu.pk/sfms/Reports/SemesterFeeCategory/FeeReconciliationReport_for_Selected_Classallbank.asp"
# REPORT_URL="https://cms.must.edu.pk/sfms/Reports/SemesterFeeCategory/SelectCriteriaAllbank.asp?wlink=FeeReconciliationReport_for_Selected_Classallbank.asp"

CREDENTIALS = {
    "username": "FEE",          
    "password": "Eef@tomust"    
}
# "90-Spring 2025"
# "87-FALL 2024"
# selected_class = "ALL"   
# session = "1"
# status= "1"
# fee_scheme="1"
# bank="1"

start_date = "1/1/2025"  # Try dd/mm/yyyy first
end_date = "2/1/2025"             

def main():
   
    with requests.Session() as session:
       
        print("Attempting login...")
        login_response = session.post(LOGIN_URL, data=CREDENTIALS)
        
        if "Login Failed" in login_response.text or login_response.url == LOGIN_URL:
            print("Login failed. Check credentials or network connection.")
            return
        
        print("Login successful")
        
      
        print("Loading report page...")
        report_response = session.get(REPORT_URL,timeout=240)
        report_soup = BeautifulSoup(report_response.text, 'html.parser')
        
       
        form_data = {
            "SelectClass": "ALL",
            "SelectSession": "-1",
            "AcademicStatus": "=1",
            "SelectScheme": "1",
            "selectBank": "-1",
            "StartDate": "01/01/2025",  # Smaller date range
            "EndDate": "31/01/2025",
            "Submit": "Continue"
        }
      
        print(" Fetching fee data...")
        data_response = session.post(REPORT_URL, data=form_data)
        data_soup = BeautifulSoup(data_response.text, 'html.parser')
        print(data_soup)
       
        data_table = data_soup.find('table', id='AutoNumber2')
        
        if not data_table:
            print("Data table not found. Possible issues:")
            print("- Semester/Scheme values need updating")
            print("- Page structure changed")
            with open('error_page.html', 'w', encoding='utf-8') as f:
                f.write(data_response.text)
            print("Saved error page to 'error_page.html'")
            return  
        print("Data table found")
        
        
        headers = [th.get_text(strip=True) for th in data_table.find('tr').find_all('td')]
        print(headers)
      
        # science_departments=["Biotechnology","Chemistry","Computer Sciences & Information Technology","Forestry","Mathematics","Physics","Physics (Pallandri Campus)","Statistics","Zoology"]
        # engineering_departments=["Civil Engineering","Civil Engineering Technology","Computer Systems Engineering","Electrical Engineering","Electrical Engineering Technology","Mirpur Institute of Technology","Mechanical Engineering","Software Engineering"]
        # social_sciences_departments =["Education","Education (Pallandri Campus)","English","Home Economics","Institute of Islamic Studies","International Relations","International Relations (Pallandri Campus)","Law","Mass Communication","Sociology","Tourism and Hospitality"]
        # health_medical_departments =["Allied Health Sciences","Human Nutrition & Dietetics","Microbiology","Pharmacy","Physiotherapy"]
        # must_businessSchool_departments=["MUST business School"]
        
        # faculties={}
        # faculty_sciences={}
        # faculty_Engineering={}
        # faculty_social_sciences={}
        # faculty_health_medical={}
        # faculty_must_business={}

        # for row in data_table.find_all('tr')[1:]:  # Skip header row
            
        #     if row.find_all('td')[0].get_text(strip=True) in science_departments:
        #         faculty_sciences[row.find_all('td')[0].get_text(strip=True)]={headers[4]:row.find_all('td')[4].get_text(strip=True),headers[11]:row.find_all('td')[11].get_text(strip=True),headers[12]:row.find_all('td')[12].get_text(strip=True)}
        #         faculties["Faculty of Natural and Applied Sciences"]=faculty_sciences
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




if __name__ == "__main__":
    main()
