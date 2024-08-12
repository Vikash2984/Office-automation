This generator requires 3 place holders from the user {name} {coure} and {date}

i) Which needs to be given as a list of dictionaries in ./app.py :
student_data = [
    {'name': 'Vikash Kumar Pandey', 'course': 'Machine Learning', 'date': '12-08-2024'},
    {'name': 'Vivek', 'course': 'Data Science', 'date': '11-08-2024'},
]
ii) .xlsx file in ./XLSS-Generator
iii) or .csv file ./CSV-Generator

1) Certificates are generated in docx format using <Generator-word> and saved in the <docx> directory
2) These docx certificates are then converted to pdf using <Converter-pdf> 
3) The final result after running the script twice is achived at ./Certificates directory