# Exchange Scripts
These scripts can help an exchange admin to do thier job better and faster.
I will update this repository every time I create a script that can be used from other admins.

Have fun with this repository

Arne Tiedemann
arne.tiedemann@tiedemanns.info


In this repository are the following scripts:

## Maintenance

### Remove-EmailAddresses
.DESCRIPTION
  This script remove unwanted proxyAddresses/EmailAddresses from any mailbox that has an proxyaddress matched by the given domainname.

.PARAMETER DomainName
  Defines the DomainName that should be removed from all Mailboxes

.PARAMETER Groups
  [Boolean]
  Define if EmailAddresses from Distributiongroups should also removed
  Value True | False

.PARAMETER PublicFolder
  [Boolean]
  Define if EmailAddresses from Email enabled PublicFolders should also removed
  Value True | False

.PARAMETER Report
  [Boolean]
  Set this Parameter to False if you want to realy remove Emailaddresses from defined objects
  Value True | False