AD-to-SWN is a modification of a closed-source python script used to query an Active Directory domain controller for a set of users and write a valid XML document ready for upload directly via SFTP to the Send Word Now alerting service.

###Status/To Do 
This is still in alpha; the closed-source script is a modification of this in current usage, and required fields are not yet abstracted out into the creation of the SWN Contacts. 

The SWN Contact class needs to be separated out and take in a listing of arguments / be able to translate those into all possible SWN fields or add them to custom fields as necessary. The XML writer will then need to be updated accordingly. Similarly, BatchProcessingDirectives need to be declared in the INI file and handled appropriately (this should be a fast fix).

Thorough integration with other options for python3-ldap has not been explored, as this was written testing in a specific environment. Handling of SSL and TLS settings needs to be implemented.


###Configuration File
Configuration information is stored in the file *swn_config.ini* in the below format. An example configuration file is included.

```INI
[LDAP Server Values]
HOST = # LDAP Server IP or DNS
PORT = # LDAP Port number, python3-ldap defaults to 389
USE_SSL = # Boolean value, python3-ldap defaults to False
ALLOWED_REFERRAL_HOSTS =  # See python3-ldap, defaults to none
TLS = None # See python3-ldap, defaults to none. Not yet implemented.

[LDAP Connection Values]
Prompt_For_Credentials = False
# Defaults to prompting for and authenticating with
# user's provided credentials in Active Directory
# Otherwise, override these values
USER = # domain\username
PASSWORD = # password

[LDAP Search Values]
SEARCH_BASE = # Search Base of LDAQ Query
SEARCH_FILTER = # Valid LDAP Query
ATTRIBUTES = # Valid AD attributes separated by ", "
PAGED_SIZE = 50

[SWN Values]
accountID = # AccountID used for Send Word Now
```
