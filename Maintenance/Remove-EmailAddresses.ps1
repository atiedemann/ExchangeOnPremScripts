[CmdletBinding()]
param
(
  [Parameter(Position=0, Mandatory=$true)]
  [string]
  $DomainName,

  [Parameter(Position=1, Mandatory=$false)]
  [bool]
  $PublicFolder,

  [Parameter(Position=2, Mandatory=$false)]
  [bool]
  $Groups,

  [Parameter(Position=2, Mandatory=$false)]
  [bool]
  $Report = $true
)
<#
.SYNOPSIS
  This script remove unwanted proxyAddresses/EmailAddresses
.DESCRIPTION
  This script remove unwanted proxyAddresses/EmailAddresses

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

.EXAMPLE

.LINK
  http://www.infowan.de
  Author: 		Arne Tiedemann
  Email: 			Arne.Tiedemann@tiedemanns.info
#>
#************************************************************************************
#		Define variables
#************************************************************************************


$DateLog = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$PathLogfile = ('{0}\Documents\infoWAN-ProxyAddresses_Remove_{1}.log' -f $env:PUBLIC, $DateLog)
$PathLogReDo = ('{0}\Documents\infoWAN-ProxyAddresses_ReDo_{1}.log' -f $env:PUBLIC, $DateLog)
$ProxySearch = "*@$($DomainName)"
If ($Report) { $MsgReport = ' -- Report only -- ' }

# Define Mailboxes that does not updated
$RegExUnwantedMailboxes = 'HealthMail|SM_|Admin'

#************************************************************************************
#		Define functions
#************************************************************************************
function Set-Logging
{
  Param(
    [Parameter(Mandatory=$false)]
    $Severity = 'Information',
    [Parameter(Mandatory=$true)]
    $Message
  )

  $Message = ("{0} {1} => {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm.ss'), $Severity, $Message)
  Write-Host $Message -ForegroundColor Yellow
  $Message | Out-File -FilePath $PathLogfile -Append
}


#************************************************************************************
#		The Script
#************************************************************************************

# Get Mailboxes with EmailAddresses
$Mailboxes = Get-Mailbox -Filter {(EmailAddresses -like $ProxySearch)} |
  Where-Object { $_.sAMAccountName -notMatch $RegExUnwantedMailboxes }

# Running Loop for alle Mailboxes
foreach($Mailbox in $Mailboxes) {
  $Addresses = ($Mailbox).EmailAddresses -match $DomainName

  # If EmailAdresses found write Log
  if ($Addresses.Count -ge 1) {
    Set-Logging -Message ('{0} Get User Mailbox {1} Information' -f $MsgReport, $Mailbox.DisplayName)
  }

  foreach ($Addr in $Addresses) {
    # Delete variables
    $RuntimeError = $false
    $Msg = $null

    # Try to remove addresses

    if ($Addr.IsPrimaryAddress) {
        Set-Logging -Severity 'Warning' -Message ('This address ({0}) is the primary emailaddress' -f $Addr)
      } else {
        try {
        $Msg += ('Trying to remove address ({0}) => ' -f $Addr)
        If ($Report -eq $false) {
          Set-Mailbox -identity $Mailbox -ErrorAction Stop -EmailAddresses @{remove=$Addr.ProxyAddressString} -Confirm:$false
          "Set-Mailbox -identity $($Mailbox.sAMAccountName) -EmailAddresses @{Add='$($Addr.ProxyAddressString)'}" |
            Out-File -FilePath $PathLogReDo -Append
          $Msg += "successfully"
        }
      } catch {
        $Msg += 'failed'
        $Msg += $_
        $RuntimeError = $true
      }

      # Logging
      if ($RuntimeError) {
        Set-Logging -Severity 'Error' -Message $Msg
      } else {
        Set-Logging -Message $Msg
      }
    }
  }
}

# Run trough Groups is enabled
if ($Groups) {
  'Groups should be removed'
  $DistributionGroups = Get-Distributiongroup -Filter {(EmailAddresses -like $ProxySearch)}

  # Running Loop for alle Mailboxes
  foreach($Group in $DistributionGroups) {
    $Addresses = ($Group).EmailAddresses -match $DomainName

    # If EmailAdresses found write Log
    if ($Addresses.Count -ge 1) {
      Set-Logging -Message ('{0} Get Group {1} Information' -f $MsgReport, $Group.DisplayName)
    }

    foreach ($Addr in $Addresses) {
      # Delete variables
      $RuntimeError = $false
      $Msg = $null

      # Try to remove addresses

      if ($Addr.IsPrimaryAddress) {
          Set-Logging -Severity 'Warning' -Message ('This address ({0}) is the primary emailaddress' -f $Addr)
        } else {
          try {
          $Msg += ('Trying to remove address ({0}) => ' -f $Addr)
          If ($Report -eq $false) {
            Set-DistributionGroup -identity $Group -ErrorAction Stop -EmailAddresses @{remove=$Addr.ProxyAddressString} -Confirm:$false
            "Set-DistributionGroup -identity $($Group.sAMAccountName) -EmailAddresses @{Add='$($Addr.ProxyAddressString)'}" |
              Out-File -FilePath $PathLogReDo -Append
            $Msg += "successfully"
          }
        } catch {
          $Msg += 'failed'
          $Msg += $_
          $RuntimeError = $true
        }

        # Logging
        if ($RuntimeError) {
          Set-Logging -Severity 'Error' -Message $Msg
        } else {
          Set-Logging -Message $Msg
        }
      }
    }
  }

}

# Run trough PublicFolder is enabled
if ($PublicFolder) {
  'PublicFolder should be removed'
}



#************************************************************************************
#		End
#************************************************************************************
