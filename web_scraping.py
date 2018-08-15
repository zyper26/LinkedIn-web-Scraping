import xlrd                                    #library for reading excel text
import xlsxwriter											#library for writing into excel file
from selenium import webdriver					#selenium library for operating on browser
from bs4 import BeautifulSoup					#beautifulsoup to excel the html of page
from time import sleep    						#sleep  witll stop the excecuting the code for limited time
import time
from selenium.webdriver.common.keys import Keys

book=xlrd.open_workbook('Input.xlsx')
#Basic information worksheet to be filled with this code...........

basic_information=xlsxwriter.Workbook('Output/BasicInformation1.xlsx',{'default_date_format': 'dd/mm/yy'})										#creating a basic+information worksheet
certifications=xlsxwriter.Workbook('Output/certification1.xlsx',{'default_date_format': 'dd/mm/yy'})
education=xlsxwriter.Workbook('Output/education1.xlsx',{'default_date_format': 'dd/mm/yy'})
experience=xlsxwriter.Workbook('Output/experience1.xlsx',{'default_date_format': 'dd/mm/yy'})
honors=xlsxwriter.Workbook('Output/honors1.xlsx',{'default_date_format': 'dd/mm/yy'})
organisations=xlsxwriter.Workbook('Output/organisations1.xlsx',{'default_date_format': 'dd/mm/yy'})
patents=xlsxwriter.Workbook('Output/patents1.xlsx',{'default_date_format': 'dd/mm/yy'})
projects=xlsxwriter.Workbook('Output/projects1.xlsx',{'default_date_format': 'dd/mm/yy'})
publications=xlsxwriter.Workbook('Output/publications1.xlsx',{'default_date_format': 'dd/mm/yy'})

basic_information_worksheet=basic_information.add_worksheet('Sheet_1')		#adding a sheet in basic_information worksheet
certification_worksheet=certifications.add_worksheet('Sheet_1')
education_worksheet=education.add_worksheet('Sheet_1')
experience_worksheet=experience.add_worksheet('Sheet_1')
honors_worksheet=honors.add_worksheet('Sheet_1')
organisations_worksheet=organisations.add_worksheet('Sheet_1')
patents_worksheet=patents.add_worksheet('Sheet_1')
projects_worksheet=projects.add_worksheet('Sheet_1')
publications_worksheet=publications.add_worksheet('Sheet_1')

first_sheet = book.sheet_by_index(0)  									#book and cells are for reading the input.xslx files
cells = first_sheet.col_slice(colx=0,start_rowx=1,end_rowx=3)
i=1
for cell in cells:
	basic_information_worksheet.write(i,0,cell.value)					#copying the information from input file to other files
	certification_worksheet.write(i,0,cell.value)
	education_worksheet.write(i,0,cell.value)
	experience_worksheet.write(i,0,cell.value)
	honors_worksheet.write(i,0,cell.value)
	organisations_worksheet.write(i,0,cell.value)
	patents_worksheet.write(i,0,cell.value)
	projects_worksheet.write(i,0,cell.value)
	publications_worksheet.write(i,0,cell.value)
	i+=1	
cells = first_sheet.col_slice(colx=1,start_rowx=1,end_rowx=3)
i=1
for cell in cells:
	basic_information_worksheet.write_url(i,1,cell.value)					#copying the information from input file to other files
	certification_worksheet.write_url(i,1,cell.value)
	education_worksheet.write_url(i,1,cell.value)
	experience_worksheet.write_url(i,1,cell.value)
	honors_worksheet.write_url(i,1,cell.value)
	organisations_worksheet.write_url(i,1,cell.value)
	patents_worksheet.write_url(i,1,cell.value)
	projects_worksheet.write_url(i,1,cell.value)
	publications_worksheet.write_url(i,1,cell.value)
	i+=1	

browser = webdriver.Firefox()											#opening browser
url = 'https://www.linkedin.com/uas/login?goback=&trk=hb_signin'		
browser.get(url)														#open the link in the browser
username = browser.find_element_by_xpath("//input[@id='session_key-login']")			#find where to put the email id of person
password = browser.find_element_by_xpath("//input[@id='session_password-login']")		#find where to put the password of person

username.send_keys("username")								#email id for login
password.send_keys("password")														#password for login

browser.find_element_by_name("signin").click()											#this will click on signin 
time.sleep(10)

cells = first_sheet.col_slice(colx=1,start_rowx=1,end_rowx=3)							#this will get required information from the read file 
i=1
for cell in cells:		
	urls=cell.value																		#this will get the links from the cells
	browser.get(urls)
	time.sleep(5)																	#this will open the required link to scrap the data from												 
	try:
		see_more_positions=browser.find_element_by_xpath("//button[@class='pv-profile-section__see-more-inline link']").click()
		time.sleep(5)
	except:
		pass
	try:
		see_1_more_postion=browser.find_element_by_xpath("//button[@class='pv-profile-section__see-more-inline link']").click()
		time.sleep(5)
	except:
		pass
	try:
		see_more=browser.find_element_by_xpath("//button[@class='button-tertiary-small mt4']").click()
		time.sleep(5)
	except:
		pass
	try:
		browser.find_element_by_xpath('//button[@class="pv-education-entity-accomplishments-block__expand"]').click()
		try:
			browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
			time.sleep(5)
		except:
			pass
	except:
		pass
	try:
		browser.find_element_by_xpath('//button[@class="pv-profile-section__card-action-bar artdeco-container-card-action-bar pv-skills-section__additional-skills"]').click()
		time.sleep(5)
	except:
		pass
	try:
		see_more_education=browser.find_element_by_xpath("//button[@class='pv-profile-section__see-more-inline link']").click()
		time.sleep(5)
	except:
		pass
	try:
		see_more=browser.find_element_by_xpath("//button[@class='truncate-multiline--button']").click()
		time.sleep(5)
	except:
		pass
	try:
		language=browser.find_element_by_xpath('//button[@data-control-name="accomplishments_expand_languages"]').click()
		time.sleep(5)
	except:
		pass			
	html=browser.page_source
	soup=BeautifulSoup(html,"lxml")											#parsing html 
	try:
		profile_picture=soup.find("img",{"class":"pv-top-card-section__image"})
		print(profile_picture.get("src"))
		link=profile_picture.get("src")
		basic_information_worksheet.write_url(i,2,link)
	except:
		pass
	try:
		headline=soup.find("h2",{"class":"pv-top-card-section__headline Sans-19px-black-85% mb1"})
		print(headline.string)
		basic_information_worksheet.write(i,3,headline.string)
	except:
		pass
	try:
		current_location=soup.find("h3",{"class":"pv-top-card-section__location Sans-17px-black-70% mb1 inline-block"})
		print(current_location.string)
		basic_information_worksheet.write(i,4,current_location.string)
	except:
		pass
	try:
		industry=soup.find("h3",{"class":"pv-top-card-section__company Sans-17px-black-70% mb1 inline-block"})
		print(industry.text.strip())
		basic_information_worksheet.write_string(i,5,industry.text.strip())
	except:
		pass	
	try:
		j=2
		try:
			for company_title in soup.find_all("li",{"class":"pv-profile-section__card-item pv-position-entity ember-view"}):
		   		title=company_title.findAll("h3",{"class":"Sans-17px-black-85%-semibold"})
		   		title1=title[0].text
		   		print(title1)
		   		experience_worksheet.write(i,j,title1)
		   		j+=8
		except:
			pass
		j=3
		try:
			for company_name in soup.find_all("li",{"class":"pv-profile-section__card-item pv-position-entity ember-view"}):
				name=company_name.findAll("span",{"class":"pv-entity__secondary-title"})
				name1=name[0].text
				print(name1)
				experience_worksheet.write(i,j,name1)
				j+=8
		except:
			pass
		j=4
		try:
			for company in soup.find_all("li",{"class":"pv-profile-section__card-item pv-position-entity ember-view"}):
				work_period=company.findAll("span")
				work_period1=work_period[3].text
				print(work_period1)
				experience_worksheet.write_datetime(i,j,work_period1)
				j+=8
		except:
			pass
		j=6
		try:
			for company in soup.find_all("li",{"class":"pv-profile-section__card-item pv-position-entity ember-view"}):
				work_duration=company.findAll("span",{"class":"pv-entity__bullet-item"})
				work_duration1=work_duration[0].text
				print(work_duration1)
				experience_worksheet.write(i,j,work_duration1.strip())
				j+=8
		except:
			pass
		j=7
		try:
			for company in soup.find_all("li",{"class":"pv-profile-section__card-item pv-position-entity ember-view"}):
				work_place=company.findAll("span")
				work_place1=work_place[7].text
				print(work_place1)
				experience_worksheet.write(i,j,work_place1.strip())
				j+=8
		except:
			pass
		try:
			for company in soup.find_all("li",{"class":"pv-profile-section__card-item pv-position-entity ember-view"}):
				work_period=company.findAll("span")
				work_period1=work_period[3].text
				name=company.findAll("span",{"class":"pv-entity__secondary-title"})
				name1=name[0].text
				if "present" in (str(work_period1)).lower():
					print(name1.strip())
					basic_infomation_worksheet.write(i,6,name1.strip())
				else:
					print(name1.strip())												 #this will print the previuos companies s/he worked in
					basic_infomation_worksheet.write(i,7,name1.strip())
		except:
			pass
	except:
		pass				  
	try:
		number_of_connections=soup.find("h3" ,{"class":"pv-top-card-section__connections pv-top-card-section__connections--with-separator Sans-17px-black-70% mb1 inline-block"})
		number_of_connections1=number_of_connections.find("span",{"aria-hidden":"true"})
		print(number_of_connections1.get_text().strip())
		basic_information_worksheet.write(i,9,number_of_connections1.get_text().strip())	
	except:
		pass
	try:
		j=2	
		try:		
			for school in soup.find_all("h3", {"class":"pv-entity__school-name Sans-17px-black-85%-semibold"}):
				print(school.text)
				education_worksheet.write(i,j,school.text.strip())
				basic_information_worksheet.write(i,8,school.text.strip())
				j+=8
		except:
			pass					
		j=3
		try:
			for dates in soup.findAll("p",{"class":"pv-entity__dates Sans-15px-black-70%"}):
				dates1=dates.findAll("span")
				dates2=dates1[1].text
				print(dates2.strip())
				education_worksheet.write(i,j,dates2.strip)
				j+=8
		except:
			pass	
		j=4
		try:
			for degree1 in soup.findAll("p",{"class":"pv-entity__secondary-title pv-entity__degree-name pv-entity__secondary-title Sans-15px-black-85%"}):
				degree2=degree1.findAll("span",{"class":"pv-entity__comma-item"})
				degree3=degree2[0].text
				print(degree3)
				education_worksheet.write(i,j,degree3.strip())
				j+=8			
		except:
			pass		
		j=5
		try:
			for field_of_study1 in soup.findAll("p",{"class":"pv-entity__secondary-title pv-entity__fos pv-entity__secondary-title Sans-15px-black-70%"}):
				field_of_study2=field_of_study1.findAll("span",{"class":"pv-entity__comma-item"})
				field_of_study3=field_of_study2[0].text
				print(field_of_study3)
				education_worksheet.write(i,j,field_of_study3)
				j+=8
		except:
			pass			
		j=6
		try:
			for grade1 in soup.findAll("p",{"class":"pv-entity__secondary-title pv-entity__grade pv-entity__secondary-title Sans-15px-black-70%"}):
				grade2=grade1.findAll("span",{"class":"pv-entity__comma-item"})
				grade3=grade2[0].text
				print(grade3)
				education_worksheet.write(i,j,grade3)
				j+=8
		except:
			pass
		j=7
		try:
			for activity in soup.findAll("p",{"class":"pv-entity__secondary-title Sans-15px-black-70%"}):
				activity1=activity.findAll("span",{"class":"activities-societies"})
				activity2=activity1[0].text
				print(activity2)
				education_worksheet.write(i,j,activity2)
				j+=8
		except:
			pass
		j=8
		try:
			describe=soup.findAll("li",{"class":"pv-profile-section__sortable-item pv-profile-section__section-info-item relative sortable-item ember-view"})
			for description in describe: 
				description1=description.find("div",{"class":"pv-entity__extra-details"})
				print(decription2.get_text().strip())
				education_worksheet.write(i,j,description2.get_text().strip())
				j+=8
		except:
			pass
	except:
		pass
	try:
		summary_text=soup.find("p", {"class":"pv-top-card-section__summary Sans-15px-black-55% mt5 pt5 ember-view"})
		print(summary_text.get_text().strip())
		basic_information_worksheet.write(i,10,summary_text.get_text().strip())
	except:
		pass
	try:
		skills=soup.findAll("li",{"class":"pv-skill-entity--featured pb5 pv-skill-entity relative ember-view"})	
		for skill in skills: 
			skill1=skill.find("div",{"class":"tooltip"})
			print(skill1.get_text().strip())
			basic_information_worksheet.write(i,12,skill1.get_text())	
	except:
		pass
	try:
		personal_details=soup.find("h1",{"class":"pv-top-card-section__name Sans-26px-black-85% mb1"})
		print(personal_details.get_text().strip())
		basic_information_worksheet.write(i,15,personal_details.get_text().strip())
	except:
		pass
	try:
		languages=soup.findAll("li",{"class":"pv-accomplishment-entity pv-accomplishment-entity--with-separator pv-accomplishment-entity--first pv-accomplishment-entity--last pv-accomplishment-entity--expanded pv-accomplishment-entity--narrow ember-view"})
		for language in languages:
			language1=language.find("h4",{"class":"pv-accomplishment-entity__title"})
			language2=language1.text
			print(language2)
			basic_information_worksheet.write(i,13,language2.strip())
	except:
		pass
	try:
		interest=soup.find("ul",{"class":"pv-profile-section__section-info section-info display-flex justify-flex-start overflow-hidden"})	
		for loop in interest.findAll('h4'):
			print(loop.text.strip())	
			basic_information_worksheet.write(i,14,loop.text.strip())
	except:
		pass	
	elm=browser.find_element_by_tag_name("body")
	elm.send_keys(Keys.END)
	time.sleep(5)		
	try:
		course=browser.find_element_by_xpath('//button[@data-control-name="accomplishments_expand_courses"]').click()
		time.sleep(5)
		try:
			course=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
			time.sleep(5)
			try:
				course=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-secondary-inline"]').click()
				time.sleep(5)
				try:
					course=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
					time.sleep(5)
					try:
						course=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
						time.sleep(5)
					except:
						pass
				except:
					pass	
			except:
				pass
		except:
			pass
	except:
		pass
	html=browser.page_source
	soup=BeautifulSoup(html,"lxml")	
	try:
		course=soup.findAll("ul",{"class":"pv-accomplishments-block__list pv-accomplishments-block__list--has-more"})
		for loop in course.findAll('h4'):
			print(loop.text.strip())	
			basic_information_worksheet_worksheet.write(i,11,loop.text.strip())
	except:
		pass
	try:
		certification=browser.find_element_by_xpath('//button[@data-control-name="accomplishments_expand_certifications"]').click()
		time.sleep(5)
		try:
			certification=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
			time.sleep(5)
			try:
				certification=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-secondary-inline"]').click()
				time.sleep(5)
				try:
					certification=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
					time.sleep(5)
					try:
						certification=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
						time.sleep(5)
					except:
						pass
				except:
					pass	
			except:
				pass
		except:
			pass	
	except:
		pass
	html=browser.page_source
	soup=BeautifulSoup(html,"lxml")	
	j=2
	try:
		certification=soup.find("ul",{"class":"pv-accomplishments-block__list"})	
		for certification1 in certification.findAll('h4'):
			print(certification1.text.strip())	
			certifications_worksheet.write(i,j,certification1.text.strip())
			j+=6
	except:
		pass	
	j=3
	try:
		certification=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		certi=certification.findAll("p",{"class":"Sans-15px-black-55% ml2"})
		for loop in certi:
			print(loop.text.strip())
			certifications_worksheet.write(i,j,loop.text.strip())
			j+=6
	except:
		pass
	j=5
	try:
		certification=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		certi=certification.findAll("a",{"class":"lazy-image EntityPhoto-square-2 loaded"})	
		for loop in pat:
			print(loop.get("src"))
			certifications_worksheet.write_url(i,j,loop.get("src"))
			j+=7
	except:
		pass			
	try:
		honor=browser.find_element_by_xpath('//button[@data-control-name="accomplishments_expand_honors"]').click()
		time.sleep(5)
		try:
			honor=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
			time.sleep(5)
			try:
				honor=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-secondary-inline"]').click()
				time.sleep(5)
				try:
					honor=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
					time.sleep(5)
					try:
						honor=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
						time.sleep(5)
					except:
						pass
				except:
					pass	
			except:
				pass
		except:
			pass			
	except:
		pass
	html=browser.page_source
	soup=BeautifulSoup(html,"lxml")	
	j=2
	try:
		honor=soup.find("ul",{"class":"pv-accomplishments-block__list pv-accomplishments-block__list--has-more"})
		for honor1 in honor.findAll('h4'):
			print(honor1.text.strip())	
			organisations_worksheet.write(i,j,honor1.text.strip())
			j+=5
	except:
		pass
	j=3
	try:
		honor=soup.find("ul",{"class":"pv-accomplishments-block__list pv-accomplishments-block__list--has-more"})
		honor1=honor.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop in honor1:
			print(loop.find("span",{"class":"pv-accomplishment-entity__issuer"}).text.strip())
			organisations_worksheet.write(i,j,loop.find("span",{"class":"pv-accomplishment-entity__issuer"}).text.strip())
			j+=5
	except:
		pass
	j=4	
	try:
		honor=soup.find("ul",{"class":"pv-accomplishments-block__list pv-accomplishments-block__list--has-more"})
		honor1=honor.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop1 in org:
			print(loop1.contents[1].text.strip())
			organisations_worksheet.write_datetime(i,j,loop1.contents[1].text.strip())
			j+=5
	except:
		pass	
	try:
		organisation=browser.find_element_by_xpath('//button[@data-control-name="accomplishments_expand_organizations"]').click()
		time.sleep(5)
		try:
			organisations1=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
			time.sleep(5)
			try:
				organisations1=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-secondary-inline"]').click()
				time.sleep(5)
				try:
					organisations1=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
					time.sleep(5)
				except:
					pass	
			except:
				pass
		except:
			pass		
	except:
		pass
	html=browser.page_source
	soup=BeautifulSoup(html,"lxml")	
	j=2
	try:
		organisation=soup.find("ul",{"class":"pv-accomplishments-block__list pv-accomplishments-block__list--has-more"})
		for org in organisation.findAll('h4'):
			print(org.text.strip())	
			organisations_worksheet.write(i,j,org.text.strip())
			j+=5
	except:
		pass
	j=3
	try:
		organisation=soup.find("ul",{"class":"pv-accomplishments-block__list pv-accomplishments-block__list--has-more"})
		org=organisation.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop in org:
			print(loop.find("span",{"class":"pv-accomplishment-entity__position"}).text.strip())
			organisations_worksheet.write(i,j,loop.find("span",{"class":"pv-accomplishment-entity__position"}).text.strip())
			j+=5
	except:
		pass
	j=4	
	try:
		organisation=soup.find("ul",{"class":"pv-accomplishments-block__list pv-accomplishments-block__list--has-more"})
		org=organisation.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop1 in org:
			print(loop1.contents[1].text.strip())
			organisations_worksheet.write_datetime(i,j,loop1.contents[1].text.strip())
			j+=5
	except:
		pass
	j=5	
	try:
		organisation=soup.find("ul",{"class":"pv-accomplishments-block__list pv-accomplishments-block__list--has-more"})
		org=organisation.findAll("p",{"class":"pv-accomplishment-entity__description Sans-15px-black-70%"})
		for loop1 in org:
			print(loop1.text.strip())
			organisations_worksheet.write_datetime(i,j,loop1.text.strip())
			j+=5
	except:
		pass
	try:
		patent=browser.find_element_by_xpath('//button[@data-control-name="accomplishments_expand_patents"]').click()
		time.sleep(5)
		try:
			patent=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
			time.sleep(5)
			try:
				patent=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-secondary-inline"]').click()
				time.sleep(5)
				try:
					patent=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
					time.sleep(5)
				except:
					pass	
			except:
				pass
		except:
			pass		
	except:
		pass
	html=browser.page_source
	soup=BeautifulSoup(html,"lxml")	
	j=4
	try:
		patent=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		for pat in patent.findAll('h4'):
			print(pat.text.strip())	
			patents_worksheet.write(i,j,pat.text.strip())
			j+=7
	except:
		pass	
	j=3
	try:
		patent=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		pat=patent.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop in pat:
			print(loop.find("span",{"class":"pv-accomplishment-entity__date"}).text.strip())
			patents_worksheet.write(i,j,loop.find("span",{"class":"pv-accomplishment-entity__date"}).text.strip())
			j+=7
	except:
		pass
	j=5
	try:
		patent=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		pat=patent.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop in pat:
			print(loop.find("span",{"class":"pv-accomplishment-entity__date"}).text.strip())
			patents_worksheet.write(i,j,loop.find("span",{"class":"pv-accomplishment-entity__date"}).text.strip())
			j+=7
	except:
		pass
	j=6
	try:
		patent=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		pat=patent.findAll("a",{"class":"pv-accomplishment-entity__external-source"})
		for loop in pat:
			print(loop.get("href"))
			patents_worksheet.write_url(i,j,loop.get("href"))
			j+=7
	except:
		pass
	j=2
	try:
		patent=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		pat=patent.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop in pat:
			print(loop.find("span",{"class":"pv-accomplishment-entity__issuer"}).text.strip())
			patents_worksheet.write(i,j,loop.find("span",{"class":"pv-accomplishment-entity__issuer"}).text.strip())
			j+=7
	except:
		pass
	try:
		publication=browser.find_element_by_xpath('//button[@data-control-name="accomplishments_expand_publications"]').click()
		time.sleep(5)
		try:
			publication=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
			time.sleep(5)
			try:
				publication=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-secondary-inline"]').click()
				time.sleep(5)
				try:
					publication=browser.find_element_by_xpath('//button[@class="pv-profile-section__see-more-inline link"]').click()
					time.sleep(5)
				except:
					pass	
			except:
				pass
		except:
			pass		
	except:
		pass
	html=browser.page_source
	soup=BeautifulSoup(html,"lxml")
	j=2
	try:
		publication=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		for pat in publication.findAll('h4'):
			print(pat.text.strip())	
			publications_worksheet.write(i,j,pat.text.strip())
			j+=6
	except:
		pass	
	j=3
	try:
		publication=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		pat=publication.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop in pat:
			print(loop.find("span",{"class":"pv-accomplishment-entity__date"}).text.strip())
			publications_worksheet.write(i,j,loop.find("span",{"class":"pv-accomplishment-entity__date"}).text.strip())
			j+=6
	except:
		pass
	j=5
	try:
		publication=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		pat=publication.findAll("p",{"class":"pv-accomplishment-entity__subtitle"})
		for loop in pat:
			print(loop.find("span",{"class":"pv-accomplishment-entity__publisher"}).text.strip())
			publications_worksheet.write(i,j,loop.find("span",{"class":"pv-accomplishment-entity__publisher"}).text.strip())
			j+=6
	except:
		pass
	j=4
	try:
		publication=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		pat=publication.findAll("a",{"class":"pv-accomplishment-entity__external-source"})
		for loop in pat:
			print(loop.get("href"))
			publications_worksheet.write_url(i,j,loop.get("href"))
			j+=6
	except:
		pass
	j=2
	try:
		publication=soup.find("ul",{"class":"pv-accomplishments-block__list"})
		pat=publication.findAll("p",{"class":"pv-accomplishment-entity__description Sans-15px-black-70%"})
		for loop in pat:
			print(loop.text.strip())
			publications_worksheet.write(i,j,loop.text.strip())
			j+=6
	except:
		pass				
	i+=1	


basic_information.close()					#copying the information from input file to other files
certifications.close()
education.close()
experience.close()
honors.close()
organisations.close()
patents.close()
projects.close()
publications.close()

browser.close()
