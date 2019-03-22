
 Clear-Host

# Old Servers
# $ExchangeServers = Get-ExchangeServer | ? { $_.AdminDisplayVersion -like 'Version 14.*' } | Sort-Object -Property Name

# New Servers
$ExchangeServers = Get-ExchangeServer | ? { $_.AdminDisplayVersion -like 'Version 15.*' } | Sort-Object -Property Name

# If Documentation is TRUE than we make the Output with Format-List and not with Format-Table -AutoSize -Wrap
$Documentation = $false

# If ConfigSet is TRUE the new configuration of Internal, external URLs will be set and the Authentication Mothods
$ConfigSet = $false

#Region Configuration
# Define internal and external Hostnames
$HostName = @{
    "Internal" = 'webmail.customer.corp'
    "External" = 'webmail.customer.com'
    'DNSDomainName' = 'customer.corp'
    'OutlookAnywhereExternal' = 'outlook.customer.com'
    'OutlookAnywhereInternal' = 'outlook.customer.corp'
    'AutoDiscoverInternal' = 'autodiscover.customer.corp'
    'AutoDiscoverExternal' = 'autodiscover.customer.com'
}

if (($HostName.Internal -eq '') -or ($HostName.External -eq '') -or ($HostName.DNSDomainName -eq ''))
{
    Write-Warning "Please set required Hostnames in this script to go forward!"
    BREAK
}

$Autodiscover = @{
    'InternalUrl' = ('https://{0}/Autodiscover/Autodiscover.xml' -f $HostName.AutoDiscoverInternal)
    'ExternalUrl' = ('https://{0}/autodiscover/Autodiscover.xml' -f $HostName.AutoDiscoverExternal)
    'BasicAuthentication' = $true
    'WindowsAuthentication' = $true
    'WSSecurityAuthentication' = $true
    'DigestAuthentication' = $false
    'OAuthAuthentication' = $false
}


$OutlookAnywhere = @{
    'InternalHostname' = $HostName.OutlookAnywhereInternal
    'ExternalHostname' = $HostName.OutlookAnywhereExternal
    'InternalClientAuthenticationMethod' = 'Ntlm'
    'ExternalClientAuthenticationMethod' = 'NTLM'
    'IISAuthenticationMethods' = 'Ntlm'
    'ExternalClientsRequireSsl' = $true
    'InternalClientsRequireSsl' = $true
    'SSLOffloading' = $false
}

$OutlookWebAccess = @{
    'InternalURL' = ('https://{0}/owa' -f $HostName.Internal)
    'ExternalUrl' = ('https://{0}/owa' -f $HostName.External)
    'FormsAuthentication' = $true
    'DigestAuthentication' = $false
    'OAuthAuthentication' = $false
    'BasicAuthentication' = $true
    'WindowsAuthentication' = $true
    'ExternalAuthenticationMethods' = 'Fba'
    'LogonFormat' = 'Username'
    'DefaultDomain' = (Get-ADDomain).NetBIOSName
}

$EcpVirtualDirectory = @{
    'InternalURL' = ('https://{0}/ecp' -f $HostName.Internal)
    'ExternalUrl' = ('https://{0}/ecp' -f $HostName.External)
    'FormsAuthentication' = $true
    'DigestAuthentication' = $false
    'BasicAuthentication' = $true
    'WindowsAuthentication' = $true
    'ExternalAuthenticationMethods' = 'Fba'
}

$OabVirtualDirectory = @{
    'InternalURL' = ('https://{0}/OAB' -f $HostName.Internal)
    'ExternalUrl' = ('https://{0}/OAB' -f $HostName.OutlookAnywhereExternal)
    'BasicAuthentication' = $true
    'OAuthAuthentication' = $false
    'WindowsAuthentication' = $true
}

$WebServicesVirtualDirectory = @{
    'InternalURL' = ('https://{0}/EWS/Exchange.asmx' -f $HostName.Internal)
    'ExternalUrl' = ('https://{0}/EWS/Exchange.asmx' -f $HostName.OutlookAnywhereExternal)
    'WSSecurityAuthentication' = $true
    'DigestAuthentication' = $false
    'OAuthAuthentication' = $false
    'BasicAuthentication' = $true
    'WindowsAuthentication' = $true
}

$MAPIVirtualDirectory = @{
    'InternalURL' = ('https://{0}/mapi' -f $HostName.Internal)
    'ExternalUrl' = ('https://{0}/mapi' -f $HostName.External)
    'IISAuthenticationMethods' = 'Basic','Ntlm','Negotiate'
}

$ActiveSyncVirtualDirectory = @{
    'InternalURL' = ('https://{0}/Microsoft-Server-ActiveSync' -f $HostName.Internal)
    'ExternalUrl' = ('https://{0}/Microsoft-Server-ActiveSync' -f $HostName.External)
    'BasicAuthentication' = $true
    'WindowsAuthentication' = $false
}

$ClientAccessServer = @{
    'AutoDiscoverServiceInternalUri' = ('https://{0}/Autodiscover/Autodiscover.xml' -f $HostName.AutoDiscoverInternal)
}

#Endregion


if ( -not ($ConfigSet))
{
    Write-Host "`n######################################`n#   Getting information from Server(s)`n######################################`n"
    ########################################
    # Only Configuration will be displayed
    ########################################
    $AutoD = @()
    $OA = @()
    $OWA = @()
    $ECP = @()
    $OAB = @()
    $EWS = @()
    $MAPI = @()
    $AS = @()
    $CAS = @()

    # Check if Exchange Server was selected
    if ($ExchangeServers.Count -ge 1) {


        Write-Host "Get-AutodiscoverVirtualDirectory" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {
                $AutoD += Get-AutodiscoverVirtualDirectory -Server $i.Name
            }
        }

        Write-Host "Get-OutlookAnywhere" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {
                $OA += Get-OutlookAnywhere -Server $i.Name
            }
        }

        Write-Host "Get-OWAVirtualDirectory" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {

                $OWA += Get-OWAVirtualDirectory -Server $i.Name
            }
        }

        Write-Host "Get-ECPVirtualDirectory" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {
                $ECP += Get-ECPVirtualDirectory -Server $i.Name
            }
        }

        Write-Host "Get-OABVirtualDirectory" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {
                $OAB += Get-OABVirtualDirectory -Server $i.Name
            }
        }

        Write-Host "Get-WebServicesVirtualDirectory" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {
                $EWS += Get-WebServicesVirtualDirectory -Server $i.Name
            }
        }

        Write-Host "Get-MAPIVirtualDirectory" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {
                $MAPI += Get-MAPIVirtualDirectory -Server $i.Name
            }
        }

        Write-Host "Get-ActiveSyncVirtualDirectory" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {
                $AS += Get-ActiveSyncVirtualDirectory -Server $i.Name
            }
        }

        Write-Host "Get-ClientAccessService" -ForegroundColor Green
        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {
                $CAS += Get-ClientAccessService -Identity $i.Name
            }
        }

        ######################################
        # Output
        ######################################
        Write-Host "`n######################################`n#   Generating output`n######################################`n"

        Write-Host "Get-AutodiscoverVirtualDirectory" -ForegroundColor Green
        if ($Documentation) {
            $AutoD | Select-Object Server,InternalUrl,ExternalUrl,InternalAuthenticationMethods,ExternalAuthenticationMethods
        } else {
            $AutoD | Select-Object Server,InternalUrl,ExternalUrl,InternalAuthenticationMethods,ExternalAuthenticationMethods | ft -AutoSize
        }

        Write-Host "Get-OutlookAnywhere" -ForegroundColor Green
        if ($Documentation) {
            $OA | Select-Object Server,InternalHostname,ExternalHostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod,IISAuthenticationMethods,ExternalClientsRequireSsl,InternalClientsRequireSsl,SSLOffloading
        } else {
            $OA | Select-Object Server,InternalHostname,ExternalHostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod,IISAuthenticationMethods,ExternalClientsRequireSsl,InternalClientsRequireSsl,SSLOffloading | ft -AutoSize
        }

        Write-Host "Get-OWAVirtualDirectory" -ForegroundColor Green
        if ($Documentation) {
            $OWA | Select-Object Server,InternalURL,ExternalURL,*auth*,LogonFormat,DefaultDomain
        } else {
            $OWA | Select-Object Server,InternalURL,ExternalURL,*auth*,LogonFormat,DefaultDomain | ft -AutoSize -Wrap
        }

        Write-Host "Get-ECPVirtualDirectory" -ForegroundColor Green
        if ($Documentation) {
            $ECP | Select-Object Server,InternalURL,ExternalURL,*auth*
        } else {
            $ECP | Select-Object Server,InternalURL,ExternalURL,*auth* | ft -AutoSize -Wrap
        }

        Write-Host "Get-OABVirtualDirectory" -ForegroundColor Green
        if ($Documentation) {
            $OAB | Select-Object Server,InternalURL,ExternalURL,*auth*
        } else {
            $OAB | Select-Object Server,InternalURL,ExternalURL,*auth* | ft -AutoSize -Wrap
        }

        Write-Host "Get-WebServicesVirtualDirectory" -ForegroundColor Green
        if ($Documentation) {
            $EWS | Select-Object Server,InternalURL,ExternalURL,*auth*
        } else {
            $EWS | Select-Object Server,InternalURL,ExternalURL,*auth* | ft -AutoSize -Wrap
        }

        Write-Host "Get-MAPIVirtualDirectory" -ForegroundColor Green
        if ($Documentation) {
            $MAPI | Select-Object Server,InternalURL,ExternalURL,*auth*
        } else {
            $MAPI | Select-Object Server,InternalURL,ExternalURL,*auth* | ft -AutoSize -Wrap
        }

        Write-Host "Get-ActiveSyncVirtualDirectory" -ForegroundColor Green
        if ($Documentation) {
            $AS | Select-Object Server,InternalURL,ExternalURL,*auth*
        } else {
            $AS | Select-Object Server,InternalURL,ExternalURL,*auth* | ft -AutoSize -Wrap
        }

        Write-Host "Get-ClientAccessServer" -ForegroundColor Green
        if ($Documentation) {
            $CAS | Select-Object Name,Fqdn,OutlookAnywhereEnabled,AutoDiscoverServiceInternalUri
        } else {
            $CAS | Select-Object Name,Fqdn,OutlookAnywhereEnabled,AutoDiscoverServiceInternalUri | ft -AutoSize -Wrap
        }
    } else { Write-Host 'No Exchange Servers are in the Array please check the Exchange Server query' -ForegroundColor Yellow }

} else {
    ########################################
    # Configuration will be updated
    ########################################

    Write-Host "Do you want to set the virtual Directories for these servers?"
    foreach ($i in $ExchangeServers.Name) {
        Write-Host $i -ForegroundColor Green
    }
    Write-Host "[Y] Yes or [N] No" -ForegroundColor Yellow


    $Result = Read-Host
    if ($Result -eq 'Y') {

        foreach($i in $ExchangeServers) {
            if ($i.IsClientAccessServer) {

                Get-AutodiscoverVirtualDirectory -Server $i.Name | Set-AutodiscoverVirtualDirectory `
                    -InternalUrl $Autodiscover.InternalUrl `
                    -ExternalUrl $Autodiscover.ExternalUrl `
                    -BasicAuthentication $Autodiscover.BasicAuthentication `
                    -WindowsAuthentication $Autodiscover.WindowsAuthentication `
                    -WSSecurityAuthentication $Autodiscover.WSSecurityAuthentication `
                    -DigestAuthentication $Autodiscover.DigestAuthentication `
                    -OAuthAuthentication $Autodiscover.OAuthAuthentication



                Write-Host ('Set OutlookAnywhere stiings for Server {0}' -f $i.Name)
                Get-OutlookAnywhere -Server $i.Name | Set-OutlookAnywhere `
                    -InternalHostname $OutlookAnywhere.InternalHostname `
                    -ExternalHostname $OutlookAnywhere.ExternalHostname `
                    -InternalClientAuthenticationMethod $OutlookAnywhere.InternalClientAuthenticationMethod `
                    -ExternalClientAuthenticationMethod $OutlookAnywhere.ExternalClientAuthenticationMethod `
                    -InternalClientsRequireSsl $OutlookAnywhere.InternalClientsRequireSsl `
                    -ExternalClientsRequireSsl $OutlookAnywhere.ExternalClientsRequireSsl `
                    -SSLOffloading $OutlookAnywhere.SSLOffloading

                Write-Host ('Set OWAVirtualDirectory stiings for Server {0}' -f $i.Name)
                Get-OWAVirtualDirectory -Server $i.Name | Set-OwaVirtualDirectory `
                    -InternalUrl $OutlookWebAccess.InternalURL `
                    -ExternalUrl $OutlookWebAccess.ExternalUrl `
                    -LogonFormat $OutlookWebAccess.LogonFormat `
                    -DefaultDomain $OutlookWebAccess.DefaultDomain `
                    -BasicAuthentication $OutlookWebAccess.BasicAuthentication `
                    -WindowsAuthentication $OutlookWebAccess.WindowsAuthentication `
                    -FormsAuthentication $OutlookWebAccess.FormsAuthentication `
                    -OAuthAuthentication $OutlookWebAccess.OAuthAuthentication `
                    -DigestAuthentication $OutlookWebAccess.DigestAuthentication `

                Write-Host ('Set ECPVirtualDirectory stiings for Server {0}' -f $i.Name)
                Get-ECPVirtualDirectory -Server $i.Name | Set-EcpVirtualDirectory `
                    -InternalUrl $EcpVirtualDirectory.InternalURL `
                    -ExternalUrl $EcpVirtualDirectory.ExternalUrl `
                    -FormsAuthentication $EcpVirtualDirectory.FormsAuthentication `
                    -DigestAuthentication $EcpVirtualDirectory.DigestAuthentication `
                    -BasicAuthentication $EcpVirtualDirectory.BasicAuthentication `
                    -WindowsAuthentication $EcpVirtualDirectory.WindowsAuthentication `
                    -ExternalAuthenticationMethods $EcpVirtualDirectory.ExternalAuthenticationMethods

                Write-Host ('Set OABVirtualDirectory stiings for Server {0}' -f $i.Name)
                Get-OABVirtualDirectory -Server $i.Name | Set-OabVirtualDirectory `
                    -InternalUrl $OabVirtualDirectory.InternalURL `
                    -ExternalUrl $OabVirtualDirectory.ExternalUrl `
                    -BasicAuthentication $OabVirtualDirectory.BasicAuthentication `
                    -OAuthAuthentication $OabVirtualDirectory.OAuthAuthentication `
                    -WindowsAuthentication $OabVirtualDirectory.WindowsAuthentication

                Write-Host ('Set WebServicesVirtualDirectory stiings for Server {0}' -f $i.Name)
                Get-WebServicesVirtualDirectory -Server $i.Name | Set-WebServicesVirtualDirectory `
                    -InternalUrl $WebServicesVirtualDirectory.InternalURL `
                    -ExternalUrl $WebServicesVirtualDirectory.ExternalUrl `
                    -WSSecurityAuthentication $WebServicesVirtualDirectory.WSSecurityAuthentication `
                    -DigestAuthentication $WebServicesVirtualDirectory.DigestAuthentication `
                    -OAuthAuthentication $WebServicesVirtualDirectory.OAuthAuthentication `
                    -BasicAuthentication $WebServicesVirtualDirectory.BasicAuthentication `
                    -WindowsAuthentication $WebServicesVirtualDirectory.WindowsAuthentication

                Write-Host ('Set MAPIVirtualDirectory stiings for Server {0}' -f $i.Name)
                Get-MAPIVirtualDirectory -Server $i.Name | Set-MapiVirtualDirectory `
                    -InternalUrl $MAPIVirtualDirectory.InternalURL `
                    -ExternalUrl $MAPIVirtualDirectory.ExternalUrl `
                    -IISAuthenticationMethods $MAPIVirtualDirectory.IISAuthenticationMethods


                Write-Host ('Set ActiveSyncVirtualDirectory stiings for Server {0}' -f $i.Name)
                Get-ActiveSyncVirtualDirectory -Server $i.Name | Set-ActiveSyncVirtualDirectory `
                    -InternalUrl $ActiveSyncVirtualDirectory.InternalURL `
                    -ExternalUrl $ActiveSyncVirtualDirectory.ExternalUrl `
                    -BasicAuthEnabled $ActiveSyncVirtualDirectory.BasicAuthentication `
                    -WindowsAuthEnabled $ActiveSyncVirtualDirectory.WindowsAuthentication


                Write-Host ('Set ClientAccessService stiings for Server {0}' -f $i.Name)
                Get-ClientAccessService -Identity $i.Name | Set-ClientAccessService `
                    -AutoDiscoverServiceInternalUri $ClientAccessServer.AutoDiscoverServiceInternalUri
            }
        }
    }

}