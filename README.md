# Office-automation
This generator requires <b>3</b> place holders from the user <b>{name}</b> <b>{coure}</b> and <b>{date}</b>

<ul>
  <li> Which needs to be given as a list of dictionaries in ./app.py :
student_data = [
    {'name': 'Vikash Kumar Pandey', 'course': 'Machine Learning', 'date': '12-08-2024'},
    {'name': 'Vivek', 'course': 'Data Science', 'date': '11-08-2024'},
] </li>
<li> .xlsx file in <i><b>./XLSS-Generator</b></i> </li>
<li> or .csv file <i><b>./CSV-Generator</b></i></li>
</ul>
<ol>
<li> Certificates are generated in docx format using <b>Generator-word</b> and saved in the <b>docx</b> directory </li>
<li> These docx certificates are then converted to pdf using <b>Converter-pdf</b> </li>
<li> The final result after running the script twice is achived at <b><i>./Certificates directory</b></i> </li>
</ol>
