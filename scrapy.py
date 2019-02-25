import requests
import json
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment


book = Workbook()
sheet = book.active

## Excel styling 
sheet['A1'].value = 'Titles'
sheet['B1'].value = 'Links'
sheet['C1'].value = 'Snippets'
sheet.column_dimensions['A'].width=48
sheet.column_dimensions['B'].width=80
sheet.column_dimensions['C'].width=80
sheet['A1'].font=Font(sz=14, bold=True)
sheet['B1'].font=Font(sz=14, bold=True)
sheet['C1'].font=Font(sz=14, bold=True)
sheet["A1"].alignment=Alignment(horizontal='center')
sheet["B1"].alignment=Alignment(horizontal='center')
sheet["C1"].alignment=Alignment(horizontal='center')

## URLs
urls = [
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 2"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 3"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 4"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 5"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 6"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 7"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 8"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 9"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 10"',
	'https://www.googleapis.com/customsearch/v1?alt=json&cx=012325137320036709742:n2w3cf_4abi&num=10&start=1&key=AIzaSyAKKRpIIdynaQyBAiS7btOMVlP1hTDJI8o&q= "staffing and recruiting" hyderabad "see all 11"'
]


## Parsing stage
xlsx = []
for url in urls:
	response = requests.get(url)
	content = response.text
	parsed = json.loads(content)

	## Get Title from Nested JSON data

	for i in range(10):
		title = parsed["items"][i]["title"]
		link = parsed["items"][i]["link"]
		snippet = parsed["items"][i]["snippet"]
		print ("- - TITLE - -\n", "title: ", title, "\n")
		print ("- - LINK - -\n", "link: ", link, "\n")
		print ("- - SNIPPET - -\n", "snippet: ", snippet, "\n")

		xlsx.append([title, link, snippet])
		
	## Saving into Excel file

	for i in range(len(xlsx)):
		for j in range(len(xlsx[i])):
			sheet.cell(row=i+2, column=j+1).value = xlsx[i][j]
	book.save(filename='Output.xlsx')


