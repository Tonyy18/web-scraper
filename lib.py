import requests
import json as _json
from bs4 import BeautifulSoup

def getOrganizations():
	url = "https://www.studystore.nl/api/ListApi/GetOrganizationsAsync/"
	re = get(url, json=True)
	if(re):
		return re

def getDepartments(organizationId):
	url = "https://www.studystore.nl/api/ListApi/GetDepartmentsByOrganizationIdAsync/?organizationId=" + organizationId
	re = get(url, json=True)
	if(re):
		return re

def getPeriods(DepartmentId):
	url = "https://www.studystore.nl/api/ListApi/GetPeriods/?parentId=" + DepartmentId
	re = get(url, json=True)
	if(re):
		return re

def getBookLists(parentId, periodId):
	url = "https://www.studystore.nl/api/ListApi/GetBooklists/?parentId=" + parentId + "&periodId=" + periodId
	re = get(url, json=True)
	if(re):
		return re

def getGroups(periodId, data = {}):
	url = "https://www.studystore.nl/api/ListApi/GetGroups/?periodId=" + periodId
	for key in data:
		url = url + "&" + key + "=" + data[key]
	re = get(url, json=True)
	if(re):
		return re

def get(url, soup = False, json = False, xml = False):
	re = requests.get(url, headers={
		"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.108 Safari/537.36"
	})
	if(re.status_code == 200):
		if(soup == True):
			return BeautifulSoup(re.text, "html.parser")
		elif(json == True):
			return _json.loads(re.text)
		elif(xml == True):
			return ET.fromstring(re.text)
		return re.text

