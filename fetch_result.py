import requests
from bs4 import BeautifulSoup
import re
import csv
import sys

first_roll=int(input("Enter frst rollnumber :"))
last_roll=int(input("Enter last rollnumber :"))
sem_no=int(input("Enter semester :"))
filename=input("Enter filename :")
filename+=".xls"
outfile = open(filename, 'w')
res_str=""
first_run=0
counter=0

if (sem_no%2):
	url="http://wbutech.net/show-result_odd.php"
	ref_url="http://wbutech.net/result_odd.php"
else:
	url="http://wbutech.net/show-result_even.php"
	ref_url="http://wbutech.net/result_even.php"
for roll_number in range(first_roll,last_roll):
	payload={
		'semno':sem_no,
		'rectype':'1',
		'rollno':roll_number
	}
	headers={
		'Host':'wbutech.net',
		'Origin':'http://wbutech.net',
		'Referer':'http://wbutech.net/result_even.php',
		'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.130 Safari/537.36'
	}

	# POST with form-encoded data
	r = requests.post(url, data=payload,headers=headers)
	if(r.url==ref_url):
		continue

	#HTML parsing
	soup=BeautifulSoup(r.text, "html.parser")
	lblContent=soup.find(id="lblContent")
	data = []
	student_details = []
	sem_marks = []

	#parse student details
	table = lblContent.find('table')
	table_body = table.find('tbody')
	rows = table_body.find_all('tr')
	for row in rows:
	    cols = row.find_all('th')
	    cols = [ele.text.strip() for ele in cols]
	    student_details.append([ele for ele in cols if ele])
	table = table.find_next('table')

	#parse Marks
	table_body = table.find('tbody')
	rows = table_body.find_all('tr')
	for row in rows:
	    cols = row.find_all('td')
	    cols = [ele.text.strip() for ele in cols]
	    data.append([ele for ele in cols if ele])
	table = table.find_next('table')

	#parse YGPA
	table_body = table.find('tbody')
	rows = table_body.find_all('tr')
	for row in rows:
	    cols = row.find_all('td')
	    cols = [ele.text.strip() for ele in cols]
	    sem_marks.append([ele for ele in cols if ele])

	if(first_run==0):
		first_run=1
		dept=student_details[0][0]
		dept=dept.split()
		print("DEPT :"+ dept[5][1:-1] + " YEAR :"+dept[2]+" COLLEGE :"+sem_marks[4][0][22:])
		res_str=("NAME,ROLL,")
		for i in range(1,len(data)-1):
			res_str+=data[i][0]+","
		res_str+=("SGPA_ODD,SGPA_EVEN,YGPA,RESULT")
		outfile.write(res_str)
	res_str="\n"
	student_name=student_details[1][0][7:]
	student_roll=student_details[1][1][11:]
	res_str+=student_name+","+student_roll+","
	for i in range(1,len(data)-1):
		res_str+=data[i][2]+","
	sgpa_odd=sem_marks[0][0][46:]
	sgpa_even=sem_marks[1][0][26:]
	ygpa=sem_marks[2][0][8:]
	result=sem_marks[3][0][9:]
	res_str+=(sgpa_odd+","+sgpa_even+","+ygpa+","+result)
	outfile.write(res_str)
	counter+=1
	print("Written "+str(counter)+" record(s)",end="\r")
print("Written "+str(counter)+" record(s)")
outfile.close