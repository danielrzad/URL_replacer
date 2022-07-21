# URL_replacer

I wrote this script to help myself with my daily tasks at work. It's ugly, not optimised but what's most important it's working and doing its job, which is to help me process .xlsx files in which I need to validate thousands of URLs and delete/replace them based on some conditions.

Whole script consists of the two parts:
First user needs to use "check_for_duplicates_in_input_xlsx_file.py" on input file and then "URL_replacer.py" on the file created by first script.

<b>I. check_for_duplicates_in_input_xlsx_file.py</b>

<em>This file checks "input_file.xlsx" file for possible duplicates in input URLs. Sites which I'm validating very often change their domain extension (for eg. .com --> .co) or even protocol (eg. http --> https). Because of that Excel wasn't able to find duplicates and that's the reason why I've made this tool. This tool is looking for duplicates in URLs but it doesn't take full URL into consideration. Instead of that it is looking for part of the netloc, path, parameters, query, fragment.</em>

How does it works?
1) Looks for duplicates in column L starting from row no. 4.
2) Analyzes found URLs by showing them to User (script is opening them in Chrome) and asking user for final validation and deletion approval.
3) Deletes found URLs and saves file with "CheckedForDuplicates___" prefix.

<b>II. URL_replacer</b>

<em>This file is responsible for replacing/replacing while extending row into couple of rows (if needed)/deleting rows based on what is placed into "links_pairs.xlsx". 
This file needs "url_replacement_file.xlsx" with scheme like</em>

<h2>HTML Table</h2>
<table>
  <tr>
    <th></th>
    <th>Column "A"</th>
  </tr>
  <tr>
    <td>Row 1</td>
    <td>1st URL from "input_file.xlsx"</td>
  </tr>
  <tr>
    <td>Row 2</td>
    <td>1_URL to replace 1st URL from "input_file.xlsx"</td>
  </tr>
  <tr>
    <td>Row 3</td>
    <td>2_URL to replace 1st URL from "input_file.xlsx"</td>
  </tr>
  <tr>
    <td>Row 4</td>
    <td>2nd URL from "input_file.xlsx"</td>
  </tr>
  <tr>
    <td>Row 5</td>
    <td>1_URL to replace 2nd URL from "input_file.xlsx"</td>
  </tr>
  <tr>
    <td>Row 6</td>
    <td>etc.</td>
  </tr>
</table>

How does it works?
1) If URL is not present in "url_replacement_file.xlsx" script will delete given row. 
2) If URL is present in "url_replacement_file.xlsx" script will replace it with new URL and add more rows if needed based on "url_replacement_file.xlsx".


Those scripts are hardcoded mainly for my personal use, but someone with basic Python skills would be able to make some changes and adapt it for hes/her usecase. They are lacking configuration and probably it would be better to write them in pandas instead of openpyxl. 
