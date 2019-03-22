# Exchange Scripts
These scripts can help an exchange admin to do thier job better and faster.
I will update this repository every time I create a script that can be used from other admins.

Have fun with this repository

Arne Tiedemann
arne.tiedemann@tiedemanns.info


In this repository are the following scripts:

## Get-SetExchangeURLs
This script can get and set all Exchange Server Url's.
ou can filter the Exchange Servers in the script to get and/or set the urls only for that servers:

- Autodiscover
- Owa
- Outlook Anywhere
- Oab
- ActiveSync
- Ews
- Ecp
- Mapi

You can specify for each Url the authentication methods.

## Maintenance

### Remove-EmailAddresses
1. DESCRIPTION
  This script remove unwanted proxyAddresses/EmailAddresses from any mailbox that has an proxyaddress matched by the given domainname.

2. PARAMETER DomainName
  Defines the DomainName that should be removed from all Mailboxes

3. PARAMETER Groups
  [Boolean]
  Define if EmailAddresses from Distributiongroups should also removed
  Value True | False

4. PARAMETER PublicFolder
  [Boolean]
  Define if EmailAddresses from Email enabled PublicFolders should also removed
  Value True | False

5. PARAMETER Report
  [Boolean]
  Set this Parameter to False if you want to realy remove Emailaddresses from defined objects
  Value True | False