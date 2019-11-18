<#
Author:			Arne Tiedemann infoWAN Datenkommunikation GmbH
E-Mail:			Arne.Tiedemann@infowan.de
Date:			2019-11-18
Description:	This scripte verifies the DNS records of every accepted domein
                and test the connectivity to web resources.
#>
Param(
    [Parameter(Mandatory)]
    $Domains,
    $HostNames = @('autodiscover','mail','outlook','owa','webmail')
)

$DNSResult = @()

###########################################################################
# Variables
###########################################################################
$TestProtocols = @('HTTP','HTTPS')
###########################################################################
# Functions
###########################################################################

###########################################################################
# Script
###########################################################################
foreach($Domain in $Domains | Where-Object { $_ -notlike '*onmicrosoft*'} ) {
    foreach($Record in $HostNames){
        try {
            # Test each DNS record for each Domain
            Write-Host ('testing DNS name: {0}.{1} ' -f $Record, $Domain) -NoNewline
            $Result = Resolve-DnsName -Name ('{0}.{1}' -f $Record, $Domain) -DnsOnly -ErrorAction Stop

            $DNSResult += [PSCustomObject]@{
                Host = $Record
                Domain = $Domain
                Destination = $Result[0].NameHost
                Type = $Result[0].Type
                Status = 'Success'
                HTTP = ''
                HTTPS = ''
            }
            # Output results
            Write-Host 'success' -ForegroundColor Green
        } catch {

            $DNSResult += [PSCustomObject]@{
                Host = $Record
                Domain = $Domain
                Destination = ''
                Type = ''
                Status = $_.Exception.Message
                HTTP = ''
                HTTPS = ''
            }

            # Output results
            Write-Host 'failed' -ForegroundColor Yellow

        }
    }
}

# Check HTTP und HTTPS Connect
foreach($WebHost in $DNSResult | Where-Object { $_.Status -eq 'Success' }) {
    # Try to connect
    foreach($Protocol in $TestProtocols) {
        try {
            # Test each DNS record for each Domain
            Write-Host ('Testing Web access to: {0}://{1}.{2} ' -f $Protocol, $WebHost.Host, $WebHost.Domain) -NoNewline

            $Result = Invoke-WebRequest -UseBasicParsing -Uri ('{0}://{1}.{2}' -f $Protocol, $WebHost.Host, $WebHost.Domain) -ErrorAction Stop
            $WebHost.$Protocol = 'Success'

            # Output results
            Write-Host 'success' -ForegroundColor Green
        } catch {
            $WebHost.$Protocol = 'Failed'

            # Output results
            Write-Host 'failed' -ForegroundColor Yellow
        }
    }
}
###########################################################################
# Finally
###########################################################################
# Cleaning Up the workspace
$DNSResult | Format-Table -AutoSize

###########################################################################
# End
###########################################################################