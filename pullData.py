import json
import csv
import xlrd
from lxml import html
import requests


# Function for spliting description into introduction and full description
def split(string, limit):
	
	if string == None:
		return ["", ""]

	temp = string[:limit]

	i = len(temp)-1
	while i >= 0:
		if temp[i] == ".":
			c = i
			break
		i -=1

	try:
		first = string[:c+1]
		end = string[c+1:]
	except:
		first = ""
		end = string

	return [first, end]


# Changes each variable from list to string
def remList(variable):
	if len(variable) == 1:
		return variable[0]
	if len(variable) < 1:
		return ""
	try:
		if len(variable) > 1:
			return variable
	except:
		pass

def hyperlink(words, link):
	try:
		return ('=HYPERLINK("'+link+'", "'+words+'")')
	except:
		return ("")


# Pulls all Partner SFDCs
### CHANGE FILE NAME ###
wb = xlrd.open_workbook("SFDC Report.xlsx")
shmain = wb.sheet_by_index(0)
wb2 = xlrd.open_workbook("Names.xlsx")
shname = wb2.sheet_by_index(0)


partnerIDsDict = {}
nameMatching = {}
for i in range(1, shname.nrows):
	row_values = shname.row_values(i)
	if len(row_values[1]) > 0:
		try:
			data = nameMatching[row_values[0]]
			data.append(row_values[1])
		except:
			data = [row_values[1]]
	else:
		data = []
	nameMatching[row_values[0]] = data
	partnerIDsDict[row_values[3]] = row_values[4]
nameMatching[""]=[]


# Stores Partner SFDCs in a dict
for i in range(11, shmain.nrows-2):
	row_values = shmain.row_values(i)

	if len(row_values[1]) > 0 and row_values[2] != "Sum" and row_values[2] != "Count":
		
		if len(nameMatching[row_values[1]]) > 0:
			for j in nameMatching[row_values[1]]:
				pname = j
				partnerIDsDict[j] = row_values[3]

# Stores all Integration SFDCs in a dict
integrationIDsDict = {}
usedints=[]
for i in range(11, shmain.nrows-2):
	row_values = shmain.row_values(i)

	if len(row_values[1]) > 0 and row_values[2] != "Sum" and row_values[2] != "Count":

		if len(nameMatching[row_values[1]]) > 0:
			pname = nameMatching[row_values[1]][0]
		else:
			pname = row_values[1]


		integrationIDsDict[pname] = {}
		templist = []


	if len(row_values[8]) > 0:
		if row_values[8] in templist:
			data = integrationIDsDict[pname][row_values[8]]
			data.append(row_values[6])
		else:
			data = [row_values[6]]
			templist.append(row_values[8])
		integrationIDsDict[pname][row_values[8]] = data


# Gets link to all partner pages and stores it in a list
mainPage = requests.get('https://www.hashicorp.com/integrations/')
mainTree = html.fromstring(mainPage.content)
partnersLink = mainTree.xpath("//a[@class='g-integration-card']/@href ")


# Create initial variables
partners = {}
integrations = {}
names = []


# Runs through each link
for i in range(len(partnersLink)):

	# Goes to the partner page and pulls all the data
	url = 'https://www.hashicorp.com' + partnersLink[i]
	print(str(i)+" | "+url)
	tempPage = requests.get(url)
	tempTree = html.fromstring(tempPage.content)

	# Pulls the desired data from page an stores it
	### XPaths must be changed if site format changes ###
	name = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[1]/h1/text()")
	product = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[1]/div/span/text()")
	description = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[1]/p/text()")
	picture = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[2]/div/img/@src")
	gettingStarted = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[2]/p/text()")
	website = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[2]/section/a[1]/@href")
	docs = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[2]/section/a[2]/@href")
	relatedResources1 = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[1]/text()")
	relatedResources1Link = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[1]/@href")
	relatedResources2 = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[2]/text()")
	relatedResources2Link = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[2]/@href")
	relatedResources3 = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[3]/text()")
	relatedResources3Link = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[3]/@href")
	relatedResources4 = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[4]/text()")
	relatedResources4Link = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[4]/@href")
	relatedResources5 = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[5]/text()")
	relatedResources5Link = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[1]/section[3]/section/a[5]/@href")
	worksWith = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[2]/section/section/a/span/text()")

	# Changes each variable from list to string
	name = remList(name)
	product = remList(product)
	description = remList(description)
	gettingStarted = remList(gettingStarted)
	website = remList(website)
	docs = remList(docs)
	relatedResources1 = remList(relatedResources1)
	relatedResources1Link = remList(relatedResources1Link)
	relatedResources2 = remList(relatedResources2)
	relatedResources2Link = remList(relatedResources2Link)
	relatedResources3 = remList(relatedResources3)
	relatedResources3Link = remList(relatedResources3Link)
	relatedResources4 = remList(relatedResources4)
	relatedResources4Link = remList(relatedResources4Link)
	relatedResources5 = remList(relatedResources5)
	relatedResources5Link = remList(relatedResources5Link)

	link1 = hyperlink(relatedResources1, relatedResources1Link)
	link2 = hyperlink(relatedResources2, relatedResources2Link)
	link3 = hyperlink(relatedResources3, relatedResources3Link)
	link4 = hyperlink(relatedResources4, relatedResources4Link)
	link5 = hyperlink(relatedResources5, relatedResources5Link)



	if len(picture) < 1:
		# Some pictures are under a different XPath
		try:
			picture = tempTree.xpath("/html/body/div/div[2]/section/section[2]/section[2]/div/picture/img/@src")
		except:
			picture = ""
	if len(picture) == 1:
		picture = picture[0]

	if len(worksWith) > 1:
		ans = ""
		for j in worksWith:
			ans += j + ", "
		worksWith = ans[:-2]
	if len(worksWith) == 1:
		worksWith = worksWith[0]
	if len(worksWith) < 1:
		worksWith = ""
	
	# Download all pics
	#r = requests.get(picture, allow_redirects=True)
	#open('pics/'+name+'.svg', 'wb').write(r.content)


	# Gets the SFDC Partner ID for the current Partner if applicable
	try:
		partnerID = partnerIDsDict[name]
	except:
		partnerID = ""

	# Gets the SFDC Integration ID for the current Integration if applicable
	try:	
		iproduct = []
		for ikeys in integrationIDsDict[name].keys():
			if product in ikeys:
				iproduct.append(ikeys)
	except:
		iproduct = [product]


	# Stores Partner info in JSON format
	partners[name]={
		'SFDC Partner ID': partnerID,
		'<title>':"",
		'<meta name="description">':"",
		'<meta property="og:image">':"",
		'Partner name': name,

		'Corporate introduction': split(description, 200)[0],
		'Full description': split(description, 200)[1],

		'Partner logo': picture,
		'Has HashiCorp Website Presence?': "",
		'If yes, add link': website,
		
		'Badges': "",
		'Partner tier': "Select",
		'Integration type': "" 

	}

	for j in iproduct:
		try:
			integrationID = integrationIDsDict[name][j]
		except:
			integrationID = [""]


		for k in range(len(integrationID)):
			usedints.append(integrationID[k])
			# Stores Integration info in JSON format 
			integrations[name + " | " + j + " | " + integrationID[k]]={
				'SFDC Integration ID': integrationID[k],
				'SFDC Partner ID': partnerID,
				'Product Integration': j,
				'Integration Name': "",

				'Introduction': split(gettingStarted, 200)[0],
				'Full description': split(gettingStarted, 200)[1],

				'Getting started CTA': docs,
				'Link 1': link1,
				'Link 2': link2,
				'Link 3': link3,
				'Link 4': link4,
				'Link 5': link5

				# Seperate URL and link name
				# 'Related Resources 1': link1,
				# 'Related Resources 1 Link': relatedResources1Link,
				# 'Related Resources 2': relatedResources2,
				# 'Related Resources 2 Link':relatedResources2Link,
				# 'Related Resources 3': relatedResources3,
				# 'Related Resources 3 Link': relatedResources3Link,
				# 'Related Resources 4': relatedResources4,
				# 'Related Resources 4 Link': relatedResources4Link,
				# 'Related Resources 5': relatedResources5,
				# 'Related Resources 5 Link': relatedResources5Link
			}
			templist=[]
			if "HCP" in j or "Terraform Cloud" in j:
				if partners[name]['Badges']== "":
					partners[name]['Badges'] = j
				else:
					templist.append(partners[name]['Badges'])
					templist.append(j)
					partners[name]['Badges'] = templist

			if partners[name]['Badges'] != "":
				partners[name]['Partner tier'] = "Premier"



for i in range(11, shmain.nrows-2):
	row_values = shmain.row_values(i)

	if len(row_values[1]) > 0:
		name = row_values[1]

	if row_values[6] not in usedints and len(row_values[6]) > 0:
		integrationID = row_values[6]
		pi = row_values[8]
		try:
			pID = row_values[3]
		except:
			pID = ""


		integrations[str(i) + " | " + j]={
			'SFDC Integration ID': integrationID,
			'SFDC Partner ID': pID,
			'Product Integration': pi,
		}



# Code if you want to create two JSON files with the data

# with open('partners.json', 'w') as outfile:
# 	outfile.write(json.dumps(partners, indent=4))

# with open('integrations.json', 'w') as outfile:
# 	outfile.write(json.dumps(integrations, indent=4))


### JSON to CSV ###
def jsonToCsv(jsonfile,csvfile):

	# Creates new CSV file
	data_file = open(csvfile, 'w')
	csv_writer = csv.writer(data_file)

	start = True
	for i in jsonfile:

		# Creates headers based on JSON keys
		if start:
			header = jsonfile[i].keys()
			csv_writer.writerow(header)
			start = False

		# Fills rows with data from JSON values
		csv_writer.writerow(jsonfile[i].values())

	data_file.close()

# Converts JSON data to CSV files
jsonToCsv(integrations,'integrations.csv')
jsonToCsv(partners,'partners.csv')