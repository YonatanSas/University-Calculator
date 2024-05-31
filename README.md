## University Admission Predictor

The University Admission Predictor is a Python script that analyzes student grades from an Excel file and predicts their chances of being accepted to a university based on their grades and the university's admission requirements.

## Features

- User Interface (GUI) for selecting the input Excel file and specifying the output file location
- Calculates the final grade for each student based on their courses and grades
- Applies bonuses to specific courses based on the number of study units
- Checks if a student has all the required courses and a minimum of 20 study units
- Generates an output Excel file with the calculated final grades and reasons for not receiving a grade (if applicable)
- Applies formatting to the output Excel file for better readability

## Prerequisites

- Python 3.x
- pandas
- openpyxl
- tkinter

## Usage

1. Input Excel file with the following columns:
   - "ת.ז" (Student ID)
   - "מקצוע" (Course)
   - "יחידות" (Study Units)
   - "ציון סופי/ מנובא" (Final/Predicted Grade)
   - "האם בוצע ניבוי" (Prediction Status)

2. Run the script:

   
   python university_admission_predictor.py
   

3. A file selection dialog will appear. Choose the input Excel file and click "OK".

4. A save file dialog will appear. Specify the location and name for the output Excel file and click "Save".

5. The script will process the data, calculate the final grades, and generate the output Excel file.

6. A success message will be displayed upon completion.


## Output

The script generates an output Excel file with the following columns:

- "תעודת זהות" (Student ID)
- "מס' יחידות לימוד" (Number of Study Units)
- "מקצועות חובה שחסרים" (Missing Required Courses)
- "ציון סופי" (Final Grade)

The final grade is calculated based on the student's courses, grades, and applied bonuses. If a student is missing required courses or has less than 20 study units, the reason will be specified in the "מקצועות חובה שחסרים" column.


## Acknowledgments

- The script utilizes the pandas and openpyxl libraries for data manipulation and Excel file handling.
- The GUI is created using the tkinter library.
