import requests
from bs4 import BeautifulSoup
import openpyxl

def extract_results(html_content):
    """Extracts hall ticket number, name, subject grades, and result from HTML content."""
    soup = BeautifulSoup(html_content, 'html.parser')
    
    hall_ticket = None
    name = None
    subject_grades = {}
    exam_fee_not_paid = False
    result = None

    # Check for "Exam Fee Not Paid" message
    if "Exam Fee Not Paid" in soup.text:
        exam_fee_not_paid = True
        print("Exam Fee Not Paid found in HTML.")
        return None, None, {}, exam_fee_not_paid, None

    # Extract Hall Ticket and Name
    personal_details_table = soup.find('table', id='AutoNumber3')
    if personal_details_table:
        rows = personal_details_table.find_all('tr')
        for row in rows:
            cells = row.find_all('td')
            if len(cells) >= 2:
                if "Hall Ticket No." in cells[0].text:
                    hall_ticket = cells[1].find('font', color="#FF0000").text.strip() if cells[1].find('font', color="#FF0000") else None
                elif "Name" in cells[0].text:
                    name = cells[1].text.strip()
    else:
        print("Personal details table not found.")

    # Extract Subject Grades
    marks_details_table = soup.find('table', id='AutoNumber4')
    if marks_details_table:
        rows = marks_details_table.find_all('tr')
        for row in rows[2:]:  # Skip header rows
            cells = row.find_all('td')
            if len(cells) >= 4:
                subject_name = cells[1].text.strip()
                grade = cells[3].text.strip()
                subject_grades[subject_name] = grade
    else:
        print("Marks details table not found.")

    #Extract Result
    result_table = soup.find('table', id='AutoNumber5')
    if result_table:
        result_rows = result_table.find_all('tr')
        if len(result_rows) > 2:
            result_cells = result_rows[2].find_all('td')
            if len(result_cells) >= 3:
                result = result_cells[2].text.strip()

    return hall_ticket, name, subject_grades, exam_fee_not_paid, result

def fetch_and_extract(htno):

    url = "https://www.osmania.ac.in/res07/20250355.jsp"
    payload = {
        "mbstatus": "SEARCH",
        "htno": htno,
        "Submit.x": 0,
        "Submit.y": 0,
    }

    try:
        response = requests.post(url, data=payload)
        response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
        print(f"Successfully fetched data for {htno}")
        return extract_results(response.text)
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for {htno}: {e}")
        return None, None, {}, True, None #Exam fee paid = true, because no response

def main():
    start_htno = 172624831001
    end_htno = 172624831005
    
    wb = openpyxl.Workbook()
    sheet = wb.active

    # Define the desired subject order
    ordered_subjects = [
        "LAW OF CONTRACT-I",
        "FAMILY LAW-I",
        "CONSTITUTIONAL LAW-I",
        "LAW OF TORTS INCL.MOTOR VEH.ACCIDENTS & CONS.PROT.LAWS",
        "ENVIRONMENTAL LAW"
    ]

    # First row headers
    header = ["Hall Ticket No.", "Name"]
    header.extend(ordered_subjects)
    header.extend(["Result", "Exam Fee Not Paid"])
    sheet.append(header)

    for htno in range(start_htno, end_htno + 1):
        htno_str = str(htno)
        hall_ticket, name, subject_grades, exam_fee_not_paid, result = fetch_and_extract(htno_str)
        
        row_data = [hall_ticket, name]
        if exam_fee_not_paid:
            row_data = [htno_str, ""] #hall ticket number and empty name.
            for _ in ordered_subjects:
                row_data.append("") #empty grades
            row_data.extend(["", True]) #empty result, exam fee not paid true
            sheet.append(row_data)
        elif hall_ticket:
            for subject in ordered_subjects:
                row_data.append(subject_grades.get(subject, ""))
            row_data.extend([result, exam_fee_not_paid])
            sheet.append(row_data)
        else:
            row_data = [htno_str, "Unknown"]
            for _ in ordered_subjects:
                row_data.append("")
            row_data.extend(["", True]) #empty result, exam fee not paid true, because no response
            sheet.append(row_data)

    wb.save("osmania_results_subjects_as_headers.xlsx")
    print("Results saved to osmania_results_subjects_as_headers.xlsx")

if __name__ == "__main__":
    main()