<#
.SYNOPSIS
This script retrieves domains from Microsoft Graph API and compares them with email domains from Halo API.

.DESCRIPTION
This script connects to the Microsoft Graph API and the Halo API to retrieve information about domains. It compares the email domains obtained from the Halo API with the domains obtained from the Microsoft Graph API. It then identifies any missing domains and stores them in a list.

.PARAMETER customerTenantId
The TenantId of the specific customer. This parameter is optional.

.INPUTS
None.

.OUTPUTS
None.

.EXAMPLE
.\Script.ps1 -customerTenantId "12345678-1234-1234-1234-1234567890ab"
This example runs the script with the specified customerTenantId.

.NOTES
You must change the name of your secrets on line 107 - 112 to match the names of the secrets in your Key Vault.
Author: Ben Weinberg
Date: 15/05/2024
Version: 1.0
#>


Param(
    # TenantId of specific customer
    [Parameter(Mandatory=$false)]
    [GUID]$customerTenantId
)

# in 7.2 the progress on Invoke-WebRequest is returned to the runbook log output
$ProgressPreference = 'SilentlyContinue'

#region ############################## Functions ####################################

function Get-MicrosoftToken {
    Param(
        # Tenant Id
        [Parameter(Mandatory=$false)]
        [guid]$TenantId,

        # Scope
        [Parameter(Mandatory=$false)]
        [string]$Scope = 'https://graph.microsoft.com/.default',

        # ApplicationID
        [Parameter(Mandatory=$true)]
        [guid]$ApplicationID,

        # ApplicationSecret
        [Parameter(Mandatory=$true)]
        [string]$ApplicationSecret,

        # RefreshToken
        [Parameter(Mandatory=$true)]
        [string]$RefreshToken
    )

    if ($TenantId) {
        $Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    }
    else {
        $Uri = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    }

    #Define the parameters for the token request
    $Body = @{
        client_id       = $ApplicationID
        client_secret   = $ApplicationSecret
        scope           = $Scope
        refresh_token   = $RefreshToken
        grant_type      = 'refresh_token'
    }

    $Params = @{
        Uri = $Uri
        Method = 'POST'
        Body = $Body
        ContentType = 'application/x-www-form-urlencoded'
        UseBasicParsing = $true
    }

    try {
        $AuthResponse = (Invoke-WebRequest @Params).Content | ConvertFrom-Json
    } catch {
        throw "Authentication Error Occured $_"
    }

    return $AuthResponse
}

#Connecting to Primes Azure Tenant to get Credentials
Connect-AzAccount | Out-Null

$Subscription = Read-Host "Enter your SAM Application Subscription ID"
$vaultname = Read-Host "Enter the name of the Key Vault containing the SAM Application Secrets"
$HaloURL =  Read-Host "Enter the URL for the Halo API in the following formatt 'name.halopsa.com'"
$SAMEmail = Read-Host "Enter the email address for the SAM Application to use for sending errors"
$recipientEmail = Read-Host "Enter the email address for the recipient of the email"

Set-AzContext -Subscription $Subscription | Out-Null

$CSPtenant = (Get-AzKeyVaultSecret -VaultName $vaultname -Name "tenantid" -AsPlainText -ErrorAction Stop)
$applicationID = (Get-AzKeyVaultSecret -VaultName $vaultname -Name "applicationid" -AsPlainText -ErrorAction Stop)
$ApplicationSecret = (Get-AzKeyVaultSecret -VaultName $vaultname -Name "applicationsecret" -AsPlainText -ErrorAction Stop)
$RefreshToken = (Get-AzKeyVaultSecret -VaultName $vaultname -Name "RefreshToken" -AsPlainText -ErrorAction Stop)
$HaloClientID = (Get-AzKeyVaultSecret -VaultName $vaultname -Name "HaloClientID" -AsPlainText -ErrorAction Stop)
$HaloClientSecret = (Get-AzKeyVaultSecret -VaultName $vaultname -Name "HaloClientSecret" -AsPlainText -ErrorAction Stop)
$baseUrl = "https://graph.microsoft.com/beta"

#endregion

$commonTokenSplat = @{
    ApplicationID = $ApplicationID
    ApplicationSecret = $ApplicationSecret
    RefreshToken = $RefreshToken
}


#Halo API Authentication
$uri = "https://$($HaloURL)/auth/token"
$Halobody = @{
    grant_type      = "client_credentials"
    client_id       = "$HaloClientID"
    client_secret   = "$HaloClientSecret"
    scope           = "all" 
}


$response = Invoke-RestMethod -Uri $uri -Method Post -Body $Halobody

if ($response) {
    $accessToken = $response.access_token
    $refreshToken = $response.refresh_token
    } else {
        Write-Error "Failed to authenticate. HTTP Status Code: $($response.StatusCode)"
        }

$Haloheader = @{
    Authorization = "Bearer $accessToken"
}

#SetURL for getting Halo Customers
$clientsuri = "https://$($HaloURL)/api/client?count=999"

#Get Halo Customers
$clients = Invoke-RestMethod -Uri $clientsuri -Method get -Headers $Haloheader

#Create Array for Missing Domains
$missingDomainList = New-Object System.Collections.ArrayList

#Loop through Each Halo Customer
foreach ($client in $clients.clients) {
    $clientinfoURL = "https://$($HaloURL)/api/client/$($client.id)?includedetails=true"
    $clientinfo = Invoke-RestMethod -Uri $clientinfoURL -Method get -Headers $Haloheader
    $AllDomains = @()
    #If Halo Customer does not have an Azure tenant ID ignore them
    if ($null -ne $clientinfo.azure_tenants.azure_tenant_id) {
        #If Halo Customer has Multiple Azure ID's then compare the count to their email domains across all sites to Office 365 domains across all tenants
        if ($clientinfo.azure_tenants.azure_tenant_id.count -gt 1) {
            $siteuri = "https://$($HaloURL)/api/site?client_id=$($client.id)"
            $sites = Invoke-RestMethod -Uri $siteuri -Method get -Headers $Haloheader
            $HaloEmailDomainsCount = foreach ($site in $sites.sites) {
                $sitedetailsuri = "https://$($HaloURL)/api/site/$($site.id)?includedetails=true"
                $siteresponse = Invoke-RestMethod -Uri $sitedetailsuri -Method get -Headers $haloheader
                if ($null -eq $siteresponse.emaildomain) {
                    continue
                }
                $siteresponse.emaildomain -split ','
            }
            $EmailDomainCount = $HaloEmailDomainsCount.count
            foreach ($azuretenant in $clientinfo.azure_tenants.azure_tenant_id) {
                Write-Output "Processing tenant: $($clientinfo.name) | $($azuretenant)"
                    try {
                        if ($token = (Get-MicrosoftToken @commonTokenSplat -TenantID $azuretenant -Scope "https://graph.microsoft.com/.default").Access_Token) {
                            $header = @{
                                Authorization = 'bearer {0}' -f $token
                                Accept        = "application/json"
                                'ConsistencyLevel'  = "eventual"
                            }
                        }
                    } catch {
                        Write-Error "Failed to authenticate to $($clientinfo.name): $($_.Exception.Message)"
                    }
            
                
                    $getDomainsLink = $baseUrl + '/domains'
                    $Domains = while ($getDomainsLink -ne $null) {
                        $getDomains = Invoke-RestMethod -Uri $getDomainsLink -Headers $header
                        $getDomainsLink = $getDomains."@odata.nextLink"
                        $getDomains.value
                        $AllDomains += $getdomains.value
                    }
                }
                    
                $filteredDomains = $AllDomains | Where-Object { $_.id -notmatch "\.onmicrosoft\.com$" -and $_.id -notmatch "\.onmicrosoftonline\.com$" -and $_.id -notmatch "\.exclaimer\.cloud$" -and $_.id -notmatch "\.ucconnect\.co\.uk$" -and $_.id -notmatch "\.excl\.cloud$" -and $_.id -notmatch "\.call2teams\.com$" -and $_.id -notmatch "\.msteams\.8x8\.com$" -and $_.id -notmatch "\.t\.via\.co\.uk$" -and $_.id -notmatch "\.nerdio\.net$"} | ForEach-Object { $_.id }
                if ($HaloEmailDomainsCount -gt 0) {
                    $missingDomains = Compare-Object -ReferenceObject $HaloEmailDomainsCount -DifferenceObject $filteredDomains
                    if ($missingDomains.count -gt 0){
                        $missingDomains | ForEach-Object {
                            if ($_.SideIndicator -eq "=>") {
                                $missingDomainList.Add([PSCustomObject]@{
                                    "ClientName" = $clientinfo.name
                                    "MissingDomain" = $_.InputObject
                                })
                            }
                        }
                    }
                } else {
                    $filteredDomains | ForEach-Object {
                        $missingDomainList.add([PSCustomObject]@{
                            "ClientName" = $clientinfo.name
                            "MissingDomain" = $_
                        })
                    }
                }
        } else {
            #For Halo Customers with only a single azure Tenant ID get a list of their domains from Office 365 and filter them
        foreach ($azuretenant in $clientinfo.azure_tenants.azure_tenant_id) {
            Write-Output "Processing tenant: $($clientinfo.name) | $($azuretenant)"
            try {
                if ($token = (Get-MicrosoftToken @commonTokenSplat -TenantID $azuretenant -Scope "https://graph.microsoft.com/.default").Access_Token) {
                    $header = @{
                        Authorization = 'bearer {0}' -f $token
                        Accept        = "application/json"
                        'ConsistencyLevel'  = "eventual"
                    }
                }
            } catch {
                Write-Error "Failed to authenticate to $($clientinfo.name): $($_.Exception.Message)"
            }

        
            $getDomainsLink = $baseUrl + '/domains'
            $Domains = while ($getDomainsLink -ne $null) {
                $getDomains = Invoke-RestMethod -Uri $getDomainsLink -Headers $header
                $getDomainsLink = $getDomains."@odata.nextLink"
                $getDomains.value
                $AllDomains += $getdomains.value
            }
            
        $filteredDomains = $AllDomains | Where-Object { $_.id -notmatch "\.onmicrosoft\.com$" -and $_.id -notmatch "\.onmicrosoftonline\.com$" -and $_.id -notmatch "\.exclaimer\.cloud$" -and $_.id -notmatch "\.ucconnect\.co\.uk$" -and $_.id -notmatch "\.excl\.cloud$" -and $_.id -notmatch "\.call2teams\.com$" -and $_.id -notmatch "\.msteams\.8x8\.com$" -and $_.id -notmatch "\.t\.via\.co\.uk$" -and $_.id -notmatch "\.nerdio\.net$"} | ForEach-Object { $_.id }
        $customerdomains = $filteredDomains -join ','

        #Update the Halo Customers Main Site with the filtered email domains
        $adddomainurl = "https://$($HaloURL)/api/site"

        $RequestBody = @{
            id            = $clientinfo.main_site_id
            emaildomain   = $customerdomains
        
        } | ConvertTo-Json -asarray -depth 10
        if ($filteredDomains -gt 0){
            try {
                $patchdomain = Invoke-RestMethod -Uri $adddomainurl -Method POST -Headers $Haloheader -Body $RequestBody -ContentType "application/json"
            } catch {
                Write-Host "Failed to update customer $($clientinfo.name) Error: $($_.Exception.Message) "
            }
        } 
    }
    }
    Start-Sleep -seconds 1
}
} 
#Auth to CSP Tenant to send email
try {
    if ($ogtoken = (Get-MicrosoftToken @commonTokenSplat -TenantID $CSPtenant -Scope "https://graph.microsoft.com/.default").Access_Token) {
        $ogheader = @{
            Authorization = 'bearer {0}' -f $ogtoken
            'Content-type'  = "application/json"
        }
    }

} catch {
    throw "Failed to authenticate to CSP tenant: $($_.Exception.Message)"
}

if ($missingDomainList.count -gt 0) {
    $htmlTable = "<table border='1'><tr><th>ClientName</th><th>MissingDomain</th></tr>"
    foreach ($item in $missingDomainList) {
        $htmlTable += "<tr><td>$($item.ClientName)</td><td>$($item.MissingDomain)</td></tr>"
    }
$htmlTable += "</table>"
    
# Set up the email parameters
$mailfrom = $SAMEmail
$recipient = $recipientEmail
$subject = "Halo Email Domain Sync"
$Emailbody = @{
ContentType = "HTML"
Content = "The Following Customers Office 365 domains is more than their site domains in halo $htmlTable"
}
$email = @{
message = @{
    toRecipients = @(
        @{
            emailAddress = @{
                address = $recipient
            }
        }
    )
    subject = $subject
    body = $Emailbody
}
saveToSentItems = "false"
}

# Build the base URL for the API call to Email customers 
$url = $baseUrl + '/users/' + $mailfrom + '/sendMail/'


# Call the REST-API to get the customer tenants
$email = Invoke-RestMethod -Method post -headers $ogheader -Uri $url -Body ($email | ConvertTo-Json -Depth 4)

}
