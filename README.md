# Office-automation
This generator requires 3 place holders from the user <b>{name}</b> <b>{coure}</b> and <b>{date}</b>

<ul>
  <li> Which needs to be given as a list of dictionaries in ./app.py :
student_data = [
    {'name': 'Vikash Kumar Pandey', 'course': 'Machine Learning', 'date': '12-08-2024'},
    {'name': 'Vivek', 'course': 'Data Science', 'date': '11-08-2024'},
] </li>
<li> .xlsx file in ./XLSS-Generator </li>
<li> or .csv file ./CSV-Generator</li>
</ul>
<ol>
<li> Certificates are generated in docx format using <Generator-word> and saved in the <docx> directory </li>
<li> These docx certificates are then converted to pdf using <Converter-pdf> </li>
<li> The final result after running the script twice is achived at ./Certificates directory </li>
</ol>
