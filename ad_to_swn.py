import re
import string
import getpass
import configparser
from sys import stdout
from datetime import datetime
from xml.dom import minidom
from xml.dom.minidom import Document

from ldap3 import Server, Connection, SEARCH_SCOPE_WHOLE_SUBTREE
from win32security import LogonUser, LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT, error

# Dictionary of States and Abbreviations valid in Send Word Now
STATES = {
        'Alaska' : 'AK',
        'Alabama' : 'AL',
        'Arkansas' : 'AR',
        'American Samoa' : 'AS',
        'Arizona' : 'AZ',
        'California' : 'CA',
        'Colorado' : 'CO',
        'Connecticut' : 'CT',
        'District of Columbia' : 'DC',
        'Delaware' : 'DE',
        'Florida' : 'FL',
        'Georgia' : 'GA',
        'Guam' : 'GU',
        'Hawaii' : 'HI',
        'Iowa' : 'IA',
        'Idaho' : 'ID',
        'Illinois' : 'IL',
        'Indiana' : 'IN',
        'Kansas' : 'KS',
        'Kentucky' : 'KY',
        'Louisiana' : 'LA',
        'Massachusetts' : 'MA',
        'Maryland' : 'MD',
        'Maine' : 'ME',
        'Michigan' : 'MI',
        'Minnesota' : 'MN',
        'Missouri' : 'MO',
        'Northern Mariana Islands' : 'MP',
        'Mississippi' : 'MS',
        'Montana' : 'MT',
        'National' : 'NA',
        'North Carolina' : 'NC',
        'North Dakota' : 'ND',
        'Nebraska' : 'NE',
        'New Hampshire' : 'NH',
        'New Jersey' : 'NJ',
        'New Mexico' : 'NM',
        'Nevada' : 'NV',
        'New York' : 'NY',
        'Ohio' : 'OH',
        'Oklahoma' : 'OK',
        'Oregon' : 'OR',
        'Pennsylvania' : 'PA',
        'Puerto Rico' : 'PR',
        'Rhode Island' : 'RI',
        'South Carolina' : 'SC',
        'South Dakota' : 'SD',
        'Tennessee' : 'TN',
        'Texas' : 'TX',
        'Utah' : 'UT',
        'Virginia' : 'VA',
        'Virgin Islands' : 'VI',
        'Vermont' : 'VT',
        'Washington' : 'WA',
        'Wisconsin' : 'WI',
        'West Virginia' : 'WV',
        'Wyoming' : 'WY'
}

# Send Word Now Contact class
# Initilization converts AD information to appropriate form
# and prepares for XML writes
class contact:

	def __init__(self, a_dict):

		self.contactFields = []
		self.contactPoints = []
		self.customContactFields = []
		self.groupList = []

		# Required Fields: uniqueID, FirstName, LastName, E-mail
		# TODO: sAMAccountName has been arbitrarily chosen for a uniqueID
		#		and should be abstracted out
		try:
			a_dict["sAMAccountName"][0].encode('utf-8')
			a_dict["sn"][0].encode('utf-8')
			a_dict["givenName"][0].encode('utf-8')
			a_dict["mail"][0].encode('utf-8')

			self.contactID = a_dict["sAMAccountName"][0]
			self.contactFields.append(("LastName", a_dict["sn"][0]))
			self.contactFields.append(("FirstName", a_dict["givenName"][0]))
			self.contactPoints.append(("Email", "Primary Email", a_dict["mail"][0]))
		except (KeyError, UnicodeError) as e:
			raise e # field is mandatory, cannot add user

		# Optional Fields
		# TODO: the list provided is partially left over from a closed-source 
		#		implementation and should be abstracted out
		for field in [
			("Address1", "streetAddress"),
			("City", "l"),
			("State", "st"),
			("PostalCode", "postalCode"),
			("Country", "co")]:
			try:
				a_dict[field[1]][0].encode('utf-8')
				self.contactFields.append(
					(field[0], a_dict[field[1]][0]))
			except (KeyError, UnicodeError):
				pass
		
		# Optional Custom Fields
		# TODO: the list provided is partially left over from a closed-source 
		#		implementation and should be abstracted out
		for field in [
			("Title", "title"),
			("Department", "department"),
			("Floor/Building", "postOfficeBox"),
			("Company", "company")]:
			try:
				a_dict[field[1]][0].encode('utf-8')
				self.customContactFields.append((
					field[0], a_dict[field[1]][0]))
			except (KeyError, UnicodeError):
				pass

		# Optional Group Fields
		# TODO: the list provided is partially left over from a closed-source 
		#		implementation and should be abstracted out
		for field in [
			"company"]:
			try:
				a_dict[field][0].encode('utf-8')
				self.groupList.append(a_dict[field][0])
			except (KeyError, UnicodeError):
				pass

		# Optional Contact Points - Voice
		# TODO: Voice contact points are not required, and have other options
		#		Office and Cell are left over from a previous implementation
		#		and will need to be expanded
		for field in [
			("Voice", "Office", "1", "telephoneNumber"),
			("Voice", "Cell", "1", "mobile")]:
			try:
				a_dict[field[3]][0].encode('utf-8')
				self.contactPoints.append((field[0], field[1], field[2], self.check_Phone_Number(a_dict[field[3]][0])))
			except (KeyError, UnicodeError):
				pass

		# Optional Contact Points - Other
		# Not Yet Implemented for Fax, SMS, etc
	
	# Helper Function to check state against SWN Dictionary
	# TODO: Extension for non-US provinces and regions
	def check_state(self, st):
		st.upper()
		if st in STATES.values():
			return st
		else: 
			st.capitalize()
		if st in STATES.keys():
			return STATES[st]
		return None

	# Helper Function to check postal code formatting (5-digits only)
	# TODO: Extension for non-US provinces and regions
	def check_Postal_Code(self, postalCode):
		r = re.compile('[%s]' % re.escape(string.punctuation))
		postalCode = (r.sub('', postalCode).replace(" ",""))
		if postalCode.isnumeric():
			if len(postalCode)==5:
				return postalCode
			elif len(postalCode)>5:
				return arg[:5]
		return None

	# Helper Function to ensure phone numbers have no country code or punctuation/spacing
	# TODO: Extension for non-Canada/US phone numbers
	def check_Phone_Number(self, telephoneNumber):
		regex = re.compile('[%s]' % re.escape(string.punctuation))
		telephoneNumber = (regex.sub('', telephoneNumber).replace(" ",""))
		if len(telephoneNumber)==11:
			return telephoneNumber[1:]
		elif len(telephoneNumber)==10:
			return telephoneNumber
		return 'NULL'

# Parses the swn_config.ini file into a dictionary of values
def ParseConfig():
	Config = configparser.ConfigParser()
	Config.read("swn_config.ini")
	sections = Config.sections()
	config_dict = {}

	for section in sections:
		options = Config.options(section)
		for option in options:
			try:
				if Config.get(section, option) == "None":
					config_dict[option] = None
				else:
					config_dict[option] = Config.get(section, option)
			except Exception as e:
				print("exception on {opt}".format(opt=option))
				config_dict[option] = None
	return config_dict

# Reads the dictionary of values and separates them into keyword argument
# dictionaries for later use
def ReadConfig(config_dict):
	if (config_dict["prompt_for_credentials"]) == 'True':
		domain, username, password = get_credentials()
		username = domain + "\\" + username
	else: 
		username = config_dict["user"]
		password = config_dict["password"]

	if (config_dict["use_ssl"]) == "True":
		ssl = True
	else:
		ssl = False

	# Used for creation of server objects in python3-ldap
	server_kwargs = {
		"host" : config_dict["host"],
		"port" : int(config_dict["port"]),
		"use_ssl" : ssl,
		"allowed_referral_hosts" : config_dict["allowed_referral_hosts"],
		"tls" : config_dict["tls"]
	}

	# Used for creation of connection objects in python3-ldap
	connection_kwargs ={
		"user" : username,
		"password" : password,
		"auto_bind" : True
	}

	# Used when calling search queries in python3-ldap
	search_kwargs = {
		"search_base" : config_dict["search_base"],
		"search_filter" : config_dict["search_filter"],
		"attributes" : config_dict["attributes"].replace(" ", "").split(sep=","),
		"paged_size" : int(config_dict["paged_size"]),
		"search_scope" : SEARCH_SCOPE_WHOLE_SUBTREE
	}

	# Incomplete list of Send Word Now related keyword arguments
	swn_kwargs = {
		"accountID" : config_dict["accountID"]
	}
	return {"server_kwargs": server_kwargs, 
			"connection_kwargs" : connection_kwargs,
			"search_kwargs" : search_kwargs,
			"swn_kwargs" : swn_kwargs}

# When called, prompt user for domain credentials and verify
# This should be used only if the script is run from the same forest/domain
# TODO: Extend to  verify from communication with AD server
def get_credentials():
	i = 0
	while (i<3):
		print("Please enter logon information:")
		domain = input("Domain: ")
		username = input("Username: ")
		password = getpass.getpass ("Password: ")
		try: 
			hUser = LogonUser(
				username, domain, password, LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT)
		except error:
			print("Invalid credentials\n")
			i += 1
		else:
			return (domain, username, password)
	raise Exception

# Queries LDAP for results using keyword argument dictionaries
# TODO: Error handling is not robust enough here
def query_LDAP(server_kwargs, connection_kwargs, search_kwargs):
	total_entries = -1
	s = Server(**server_kwargs)
	connection_kwargs["server"] = s
	c = Connection(**connection_kwargs)
	c.search(**search_kwargs)

	total_entries += len(c.response)
	stdout.write("%3d entries returned\r" % total_entries)
	response = c.response[:-1]
	cookie = c.result['controls']['1.2.840.113556.1.4.319']['value']['cookie']

	while cookie:
		search_kwargs["paged_cookie"] = cookie
		c.search(**search_kwargs)
		stdout.flush()
		stdout.write("%3d entries returned\r" % total_entries)

		for entry in c.response:
			total_entries += 1
			response.append(entry)
		cookie = c.result['controls']['1.2.840.113556.1.4.319']['value']['cookie']

	print("Total response length: ", total_entries)
	print("\n")
	return response

# Writes XMl using minidom according to current SWN standards
def write_xml(contact_list, file_name, accountID):

	doc = Document()
	
	# Write Header
	batch = doc.createElement("contactBatch")
	batch.setAttribute("xmlns", "http://www.sendwordnow.com")
	batch.setAttribute("version", "1.0.2")

	batch_pd = doc.createElement("batchProcessingDirectives")

	batch_acc = doc.createElement("accountID")
	batch_acc.setAttribute("username", accountID)

	batch_file = doc.createElement("batchFile")
	batch_file.setAttribute("requestFileName", file_name)

	batch_dcnib = doc.createElement("batchProcessingOption")
	batch_dcnib.setAttribute("name", "DeleteContactsNotInBatch")
	batch_dcnib.setAttribute("value", "true")

	batch_pd.appendChild(batch_acc)
	batch_pd.appendChild(batch_file)
	batch_pd.appendChild(batch_dcnib)

	batch.appendChild(batch_pd)

	# Write Contacts in Contact List
	batch_cl = doc.createElement("batchContactList")

	for entry in contact_list:
		contact = doc.createElement("contact")
		contact.setAttribute("contactID", entry.contactID)
		contact.setAttribute("action", "AddOrModify")

		for duple in entry.contactFields:
			contactField = doc.createElement("contactField")
			contactField.setAttribute("name", duple[0])
			contactField.appendChild(doc.createTextNode(duple[1]))
			contact.appendChild(contactField)
		for duple in entry.customContactFields:
			contactField = doc.createElement("contactField")
			contactField.setAttribute("name", "CustomField")
			contactField.setAttribute("customName", duple[0])
			contactField.appendChild(doc.createTextNode(duple[1]))
			contact.appendChild(contactField)
		groupList = doc.createElement("groupList")
		for group in entry.groupList:
			groupName = doc.createElement("groupName")
			groupName.appendChild(doc.createTextNode(group))
			groupList.appendChild(groupName)
		contact.appendChild(groupList)

		contactPointList = doc.createElement("contactPointList")
		for cP in entry.contactPoints:
			if len(cP) == 4:
				contactPoint = doc.createElement("contactPoint")
				contactPoint.setAttribute("type", cP[0])

				contactPointField = doc.createElement("contactPointField")
				contactPointField.setAttribute("name", "Label")
				contactPointField.appendChild(doc.createTextNode(cP[1]))
				contactPoint.appendChild(contactPointField)

				contactPointField = doc.createElement("contactPointField")
				contactPointField.setAttribute("name", "CountryCode")
				contactPointField.appendChild(doc.createTextNode(cP[2]))
				contactPoint.appendChild(contactPointField)

				contactPointField = doc.createElement("contactPointField")
				contactPointField.setAttribute("name", "Number")
				contactPointField.appendChild(doc.createTextNode(cP[3]))
				contactPoint.appendChild(contactPointField)
				contactPointList.appendChild(contactPoint)
			if len(cP) == 3:
				contactPoint = doc.createElement("contactPoint")
				contactPoint.setAttribute("type", cP[0])

				contactPointField = doc.createElement("contactPointField")
				contactPointField.setAttribute("name", "Label")
				contactPointField.appendChild(doc.createTextNode(cP[1]))
				contactPoint.appendChild(contactPointField)

				contactPointField = doc.createElement("contactPointField")
				contactPointField.setAttribute("name", "Address")
				contactPointField.appendChild(doc.createTextNode(cP[2]))
				contactPoint.appendChild(contactPointField)
				contactPointList.appendChild(contactPoint)
		contact.appendChild(contactPointList)
		batch_cl.appendChild(contact)
	batch.appendChild(batch_cl)
	doc.appendChild(batch)
	
	doc.writexml(open("writing_"+file_name, 'w', encoding='utf-8'),
				 indent="	",
				 addindent="	",
				 newl="\n",
				 encoding="UTF-8")
	doc.unlink()

# Create a file name per SWN standards, ensuring it is in increasing order due to timestamp
def get_File_Name():
	now = datetime.now()
	return "request_{number}.xml".format(
		number=now.strftime("%Y%m%d%H%M%S"))

def main():
	config_dict = ParseConfig()
	kwargs_list = ReadConfig(config_dict)
	LDAP_results = query_LDAP(**kwargs_list)
	contact_list = []

	for entry in LDAP_results:
		try:
			contact_list += [contact(entry["attributes"])]
		except KeyError as e:
			pass

	file_name = get_File_Name()
	write_xml(contact_list, file_name, kwargs_list["swn_kwargs"]["accountID"])


if __name__=="__main__":
	main()