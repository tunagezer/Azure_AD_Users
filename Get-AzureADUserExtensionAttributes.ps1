function Get-AzureADUserExtensionAttributes {

    ############################################################################

    <#
    .SYNOPSIS

        Gets Azure Active Directory users all 15 extension attributes.


    .DESCRIPTION

        Gets Azure Active Directory users extension attributes in details
        using the onPremisesExtensionAttributes.extensionAttribute1 to 15 attributes.

            Use -All to get details for all users in the target tenant.

            Use -UserObjectId to target a single user or groups of users.

            Use -GuestInfo to include additional information specific to guest accounts

        Can also produce a date and time stamped CSV file as output.
        


        PRE-REQUISITE - the function uses the MSAL.ps module from the PS Gallery:
        
                        https://www.powershellgallery.com/packages/MSAL.ps


    .EXAMPLE

        Get-AzureADUserExtensionAttributes -TenantId c666a536-cb76-4360-a8bb-6593cf4d9c7f -All

        Gets the extension attributes for all users on the tenant.


    .EXAMPLE

        Get-AzureADUserExtensionAttributes -TenantId c666a536-cb76-4360-a8bb-6593cf4d9c7f 
        -UserObjectId 69447235-0974-4af6-bfa3-d0e922a92048 -CsvOutput

        Gets the extension attributes for the user, targeted by their object ID or UPN.

        Writes the output to a date and time stamped CSV file in the execution directory.


    .EXAMPLE

        Get-AzureADUserExtensionAttributes -TenantId c666a536-cb76-4360-a8bb-6593cf4d9c7f 
        -GuestInfo -CsvOutput

        Gets all users extension attributes. 

        Writes the output to a date and time stamped CSV file in the execution directory.
        
        Includes additional attributes for guest user insight.


    .EXAMPLE

        Get-AzureADUserExtensionAttributes -TenantId c666a536-cb76-4360-a8bb-6593cf4d9c7f
        -All -GuestInfo

        Gets all users extension attributes. 

        Includes additional attributes for guest user insight.


    #>

    ############################################################################

    [CmdletBinding()]
    param(

        #The tenant ID
        [Parameter(Mandatory,Position=0)]
        [string]$TenantId,

        #Get extension attributes for all users in the tenant
        [Parameter(Mandatory,Position=1,ParameterSetName="All")]
        [switch]$All,

        #Get the extension attributes for a single user by object ID or UPN
        [Parameter(Mandatory,Position=2,ParameterSetName="UserObjectId")]
        [string]$UserObjectId,

        #Include additional information for guest accounts
        [Parameter(Position=3)]
        [switch]$GuestInfo,

        #Use this switch to create a date and time stamped CSV file
        [Parameter(Position=4)]
        [switch]$CsvOutput

    )


    ############################################################################

    ##################
    ##################
    #region FUNCTIONS

    function Get-AzureADApiToken {

        ############################################################################

        <#
        .SYNOPSIS

            Get an access token for use with the API cmdlets.


        .DESCRIPTION

            Uses MSAL.ps to obtain an access token. Has an option to refresh an existing token.

        .EXAMPLE

            Get-AzureADApiToken -TenantId c666a536-cb76-4360-a8bb-6593cf4d9c7f

            Gets or refreshes an access token for making API calls for the tenant ID
            c666a536-cb76-4360-a8bb-6593cf4d9c7f.


        .EXAMPLE

            Get-AzureADApiToken -TenantId c666a536-cb76-4360-a8bb-6593cf4d9c7f -ForceRefresh

            Gets or refreshes an access token for making API calls for the tenant ID
            c666a536-cb76-4360-a8bb-6593cf4d9c7f.

        #>

        ############################################################################

        [CmdletBinding()]
        param(

            #The tenant ID
            [Parameter(Mandatory,Position=0)]
            [guid]$TenantId,

            #The tenant ID
            [Parameter(Position=1)]
            [switch]$ForceRefresh

        )


        ############################################################################


        #Get an access token using the PowerShell client ID
        $ClientId = "1b730954-1685-4b74-9bfd-dac224a7b894"
        $RedirectUri = "urn:ietf:wg:oauth:2.0:oob"
        $Authority = "https://login.microsoftonline.com/$TenantId"
    
        if ($ForceRefresh) {

            Write-Verbose -Message "$(Get-Date -f T) - Attempting to refresh an existing access token"

            #Attempt to refresh access token
            try {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -ForceRefresh
            }
            catch {}

            #Error handling for token acquisition
            if ($Response) {

                Write-Verbose -Message "$(Get-Date -f T) - API Access Token refreshed - new expiry: $(($Response).ExpiresOn.UtcDateTime)"

                return $Response

            }
            else {
            
                Write-Warning -Message "$(Get-Date -f T) - Failed to refresh Access Token - try re-running the cmdlet again"

            }

        }
        else {

            Write-Verbose -Message "$(Get-Date -f T) - Checking token cache"

            #Run this to obtain an access token - should prompt on first run to select the account used for future operations
            try {

                $Response = Get-MsalToken -ClientId $ClientId -RedirectUri $RedirectUri -Authority $Authority -Prompt SelectAccount
            }
            catch {}

            #Error handling for token acquisition
            if ($Response) {

                Write-Verbose -Message "$(Get-Date -f T) - API Access Token obtained"

                return $Response

            }
            else {

                Write-Warning -Message "$(Get-Date -f T) - Failed to obtain an Access Token - try re-running the cmdlet again"
                Write-Warning -Message "$(Get-Date -f T) - If the problem persists, start a new PowerShell session"

            }

        }


    }   #end function


    function Get-AzureADHeader {
    
        param($Token)

        return @{

            "Authorization" = ("Bearer {0}" -f $Token);
            "Content-Type" = "application/json";

        }

    }   #end function

    #endregion

    ############################################################################

    ##################
    ##################
    #region MAIN

    #Deal with different search criterea
    if ($All) {

        #API endpoint
        $Filter = "?`$select=displayName,userPrincipalName,Id,onPremisesExtensionAttributes,userType,externalUserState,creationType,createdDateTime"

        Write-Verbose -Message "$(Get-Date -f T) - All user mode selected"

    }
    elseif ($UserObjectId) {

        #API endpoint
        $Filter = "?`$filter=ID eq '$UserObjectId'&`$select=displayName,userPrincipalName,Id,onPremisesExtensionAttributes,userType,externalUserState,creationType,createdDateTime"

        Write-Verbose -Message "$(Get-Date -f T) - Single user mode selected"

    }
#
#
#    ############################################################################
    
    $Url = "https://graph.microsoft.com/v1.0/users$Filter"


    ############################################################################

    #Get / refresh an access token
    $Token = (Get-AzureADApiToken -TenantId $TenantId).AccessToken

    if ($Token) {

        #Construct header with access token
        $Headers = Get-AzureADHeader($Token)

        #Tracking variables
        $Count = 0
        $RetryCount = 0
        $OneSuccessfulFetch = $false
        $TotalReport = $null


        #Do until the fetch URL is null
        do {

            Write-Verbose -Message "$(Get-Date -f T) - Invoking web request for $Url"

            ##################################
            #Do our stuff with error handling
            try {

                #Invoke the web request
                $MyReport = (Invoke-WebRequest -UseBasicParsing -Headers $Headers -Uri $Url -Verbose:$false)

            }
            catch [System.Net.WebException] {
        
                $StatusCode = [int]$_.Exception.Response.StatusCode
                Write-Warning -Message "$(Get-Date -f T) - $($_.Exception.Message)"

                #Check what's gone wrong
                if (($StatusCode -eq 401) -and ($OneSuccessfulFetch)) {

                    #Token might have expired; renew token and try again
                    $Token = (Get-AzureADApiToken -TenantId $TenantId).AccessToken
                    $Headers = Get-AzureADHeader($Token)
                    $OneSuccessfulFetch = $False

                }
                elseif (($StatusCode -eq 429) -or ($StatusCode -eq 504) -or ($StatusCode -eq 503)) {

                    #Throttled request or a temporary issue, wait for a few seconds and retry
                    Start-Sleep -Seconds 5

                }
                elseif (($StatusCode -eq 403) -or ($StatusCode -eq 401)) {

                    Write-Warning -Message "$(Get-Date -f T) - Please check the permissions of the user"
                    break

                }
                elseif ($StatusCode -eq 400) {

                    Write-Warning -Message "$(Get-Date -f T) - Please check the query used"
                    break

                    }
                else {
            
                    #Retry up to 5 times
                    if ($RetryCount -lt 5) {
                
                        write-output "Retrying..."
                        $RetryCount++

                    }
                    else {
                
                        #Write to host and exit loop
                        Write-Warning -Message "$(Get-Date -f T) - Download request failed. Please try again in the future"
                        break

                    }

                }

            }
            catch {

                #Write error details to host
                Write-Warning -Message "$(Get-Date -f T) - $($_.Exception)"


                #Retry up to 5 times    
                if ($RetryCount -lt 5) {

                    write-output "Retrying..."
                    $RetryCount++

                }
                else {

                    #Write to host and exit loop
                    Write-Warning -Message "$(Get-Date -f T) - Download request failed - please try again in the future"
                    break

                }

            } # end try / catch


            ###############################
            #Convert the content from JSON
            $ConvertedReport = ($MyReport.Content | ConvertFrom-Json).value

            $TotalObjects = @()

            foreach ($User in $ConvertedReport) {

                if ($GuestInfo) {

                    #Construct a custom object
                    $Properties = [PSCustomObject]@{

                        displayName = $User.displayName
                        userPrincipalName = $User.userPrincipalName
                        objectId = $User.Id
                        extensionAttribute1 = $User.onPremisesExtensionAttributes.extensionAttribute1
                        extensionAttribute2 = $User.onPremisesExtensionAttributes.extensionAttribute2
                        extensionAttribute3 = $User.onPremisesExtensionAttributes.extensionAttribute3
                        extensionAttribute4 = $User.onPremisesExtensionAttributes.extensionAttribute4
                        extensionAttribute5 = $User.onPremisesExtensionAttributes.extensionAttribute5
                        extensionAttribute6 = $User.onPremisesExtensionAttributes.extensionAttribute6
                        extensionAttribute7 = $User.onPremisesExtensionAttributes.extensionAttribute7
                        extensionAttribute8 = $User.onPremisesExtensionAttributes.extensionAttribute8
                        extensionAttribute9 = $User.onPremisesExtensionAttributes.extensionAttribute9
                        extensionAttribute10 = $User.onPremisesExtensionAttributes.extensionAttribute10
                        extensionAttribute11 = $User.onPremisesExtensionAttributes.extensionAttribute11
                        extensionAttribute12 = $User.onPremisesExtensionAttributes.extensionAttribute12
                        extensionAttribute13 = $User.onPremisesExtensionAttributes.extensionAttribute13
                        extensionAttribute14 = $User.onPremisesExtensionAttributes.extensionAttribute14
                        extensionAttribute15 = $User.onPremisesExtensionAttributes.extensionAttribute15
                        userType = $User.userType
                        createdDateTime = $User.createdDateTime
                        externalUserState = $User.externalUserState
                        creationType = $User.creationType
                    

                    }
            
                }
                else {

                    #Construct a custom object
                    $Properties = [PSCustomObject]@{

                        displayName = $User.displayName
                        userPrincipalName = $User.userPrincipalName
                        objectId = $User.Id
                        extensionAttribute1 = $User.onPremisesExtensionAttributes.extensionAttribute1
                        extensionAttribute2 = $User.onPremisesExtensionAttributes.extensionAttribute2
                        extensionAttribute3 = $User.onPremisesExtensionAttributes.extensionAttribute3
                        extensionAttribute4 = $User.onPremisesExtensionAttributes.extensionAttribute4
                        extensionAttribute5 = $User.onPremisesExtensionAttributes.extensionAttribute5
                        extensionAttribute6 = $User.onPremisesExtensionAttributes.extensionAttribute6
                        extensionAttribute7 = $User.onPremisesExtensionAttributes.extensionAttribute7
                        extensionAttribute8 = $User.onPremisesExtensionAttributes.extensionAttribute8
                        extensionAttribute9 = $User.onPremisesExtensionAttributes.extensionAttribute9
                        extensionAttribute10 = $User.onPremisesExtensionAttributes.extensionAttribute10
                        extensionAttribute11 = $User.onPremisesExtensionAttributes.extensionAttribute11
                        extensionAttribute12 = $User.onPremisesExtensionAttributes.extensionAttribute12
                        extensionAttribute13 = $User.onPremisesExtensionAttributes.extensionAttribute13
                        extensionAttribute14 = $User.onPremisesExtensionAttributes.extensionAttribute14
                        extensionAttribute15 = $User.onPremisesExtensionAttributes.extensionAttribute15

                    }

                }

                $TotalObjects += $Properties
            }


            #Add to concatenated findings
            [array]$TotalReport += $TotalObjects

            #Update the fetch url to include the paging element
            $Url = ($myReport.Content | ConvertFrom-Json).'@odata.nextLink'

            #Update the access token on the second iteration
            if ($OneSuccessfulFetch) {
                
                $Token = (Get-AzureADApiToken -TenantId $TenantId).AccessToken
                $Headers = Get-AzureADHeader($Token)

            }

            #Update count and show for this cycle
            $Count = $Count + $ConvertedReport.Count
            Write-Verbose -Message "$(Get-Date -f T) - Total records fetched: $count"

            #Update tracking variables
            $OneSuccessfulFetch = $true
            $RetryCount = 0


        } while ($Url) #end do / while


    }

    #See if we need to write to CSV
    if ($CsvOutput) {

        #Output file
        $now = "{0:yyyyMMdd_hhmmss}" -f (Get-Date)
        $CsvName = "UserExtensionAttributeDetails_$now.csv"

        Write-Verbose -Message "$(Get-Date -f T) - Generating a CSV for user extension attributes"

        $TotalReport | Export-Csv -Path $CsvName -NoTypeInformation

        Write-Verbose -Message "$(Get-Date -f T) - Uuser extension attributes written to $(Get-Location)\$CsvName"

    }
    else {

        #Return stuff
        $TotalReport

    }

    #endregion


}   #end function
