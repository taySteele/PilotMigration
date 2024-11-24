<# CONFIGURATION SETTINGS #>

$env:PNPPOWERSHELL_UPDATECHECK="false"
$script:baseURL = "https://vfogonline.sharepoint.com"
$script:adminURL = "https://vfogonline-admin.sharepoint.com"
$script:clientId = "2ac4ddf7-8327-4dfb-880f-af77133fd2fa"
$script:owners = "taylor.steele@vfog.net", "BernardineAdmin@vfogonline.onmicrosoft.com", "nick.wright@vfog.net"
$script:themeName = "MPI Theme1 Green"
$script:logFileFolderPath = $PSScriptRoot + "\Logs"
$script:templateFolderPath = $PSScriptRoot + "\Templates"
$script:logoFile = $PSScriptRoot + "\Logos\siteLogo.png" 
$script:filePath = $PSScriptRoot + "\FoundationBuildProvisioning.csv" 
$script:listsFilePath = $PSScriptRoot + "\FoundationBuildProvisioning_Lists.csv" 
$script:customFields = @(
    [pscustomobject]@{InternalName='COHESION-LEGACY-ID';DisplayName='COHESION-LEGACY-ID';Type='Text';Group='*MPI Columns'}
    [pscustomobject]@{InternalName='COHESION-LEGACY-URI';DisplayName='COHESION-LEGACY-URI';Type='Note';Group='*MPI Columns'}
)
$script:majorVersions = 50000
$script:minorVersions = 2
$script:csvData = @()
$script:listCsvData = @()
$script:winauthIndex = 29

<# FUNCTIONS #>

function createLog {
    "Script started at $startTimestamp" | Out-File -FilePath $logFilePath -Force
}

function logOutput {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "HH:mm"
    $logMessage = "$timestamp - $message"
    $logMessage | Add-Content -Path $logFilePath
}

function ReadCsv ($filePath) {
    $script:csv = Import-Csv -Path $filePath
    $script:listsCsv = Import-Csv -Path $listsFilePath
}

function ProvisionSite ($provBaseTemplate, $newSiteName, $newSiteShortCode, $newSiteFullUrl) {
    logOutput "   Provisioning site $newSiteShortCode..."

    # Connect to admin center
    try {
        Connect-PnPOnline -Url $adminURL -Interactive -ClientId $clientId
    }
    catch {
        logOutput "   Unable to connect to Admin Center"
        exit
    }
    
    # Check if site exists, if not, provision one
    try {
        $existingSite = Get-PnPTenantSite -Identity $newSiteFullUrl -ErrorAction SilentlyContinue
        if (!$existingSite) {
            if ($provBaseTemplate -eq "CommunicationSite") {
                $script:newSite = New-PnPSite -Title $newSiteName -Url $newSiteFullUrl -Type $provBaseTemplate -TimeZone 17 -Wait
            }
            else {
                $script:newSite = New-PnPSite -Title $newSiteName -Alias $newSiteShortCode -Type $provBaseTemplate -TimeZone 17 -Wait
            }
            logOutput "   Finished provisioning site $newSiteShortCode."
        } 
        else {
            $script:newSite = $existingSite
            logOutput "   Site already exists for $newSiteShortCode"
        }
        # Connect to site
        Connect-PnPOnline -Url $newSite.Url -Interactive -ClientId $clientId
    } catch {
        logOutput "   Failed to provision site $newSiteShortCode. Error: $_"
        Exit
    }
}

function SiteConfiguration ($newSiteShortCode, $newSite, $template, $templatePath) {
    logOutput "   Applying configuration to site $newSiteShortCode..."
    try {
        # Add admins
        logOutput "      Adding site collection admins..."
        Add-PnPSiteCollectionAdmin -Owners $owners
        logOutput "      Added site collection admins."

        # Enable document sets
        logOutput "      Enabling document sets..."
        Enable-PnPFeature -Scope Site -Identity "3bae86a2-776d-499d-9db8-fa4cdc7884f8" -Force
        logOutput "      Enabled document sets."

        # Apply template - landing page and site settings
        logOutput "      Apply xml template..."
        Invoke-PnPSiteTemplate -ClearNavigation -Path $templatePath
        logOutput "      Applied xml template."

        # Set theme
        logOutput "      Setting theme..."
        Set-PnPWebTheme -Theme $themeName
        logOutput "      Set theme."

        # Set logo
        logOutput "      Setting logo..."
        Set-PnPSite -LogoFilePath $logoFile
        logOutput "      Setting logo."

        # Set permission settings
        #$siteGroups = Get-PnPSiteGroup
        #$membersGroup = $siteGroups | Where-Object {$_.Title -like "*Members"}
        #Set-PnPGroupPermissions -Identity $membersGroup.Title -AddRole Contribute -RemoveRole Edit

        # Set sharing settings
        logOutput "      Setting sharing settings..."
        $web = Get-PnPWeb -Includes MembersCanShare, AssociatedMemberGroup.AllowMembersEditMembership
        $web.MembersCanShare=$true
        $web.AssociatedMemberGroup.AllowMembersEditMembership=$false
        $web.AssociatedMemberGroup.Update()
        $web.RequestAccessEmail = $null
        $web.Update()
        $web.Context.ExecuteQuery()
        logOutput "      Set sharing settings."

        # Disable Workflow Task Content Type
        logOutput "      Disabling workflow task content type and three-state workflow..."
        Disable-PnPFeature -Identity 57311b7a-9afd-4ff0-866e-9393ad6647b1 -Force

        # Disable Three-state workflow
        Disable-PnPFeature -Scope Site -Identity fde5d850-671e-4143-950a-87b473922dc7 -Force
        logOutput "      Disabled workflow task content type and three-state workflow."

        # Configure Doc ID prefix
        logOutput "      Enable and configure DocID..."
        Set-PnPTenantSite -Identity $newSite -DenyAddAndCustomizePages:$false
        Set-PnPPropertyBagValue -Key "docid_msft_hier_siteprefix" -Value "MPIDOCID"
        Set-PnPTenantSite -Identity $newSite -DenyAddAndCustomizePages:$true
        logOutput "      Enabled and configured DocID."

        # Remove recent link from nav
        logOutput "      Remove Recent link from nav..."
        $navigationNodes = Get-PnPNavigationNode
        if ($navigationNodes | Where-Object {$_.Title -eq "Recent"}) {
            $recentLink = $navigationNodes | Where-Object {$_.Title -eq "Recent"}
            Remove-PnPNavigationNode -Identity $recentLink -Force
        }
        logOutput "      Removed Recent link from nav."
        logOutput "   Finished applying configuration to site $newSiteShortCode."
    } catch {
        logOutput "   Failed to apply configuration to site $newSiteShortCode. Error $_"
        Exit
    }
}

function TeamifySite ($newSite) {
    try {
        logOutput "   Associating Microsoft Team to site"
        $site = Get-PnPTenantSite -Identity $newSite
        New-PnPTeamsTeam -GroupId $site.GroupId.Guid
        logOutput "   Microsoft Team associated to site"
    } catch {
        logOutput "   Failed to connect site to Microsoft Team. Error $_"
        Exit
    }
    
}

function RegisterHub ($newSiteFullUrl, $hubName) {
    logOutput "   Registering $provShortcode as hub."
    try {
        $existingHub = Get-PnPHubSite | Where-Object {$_.Title -eq $hubName}
        if (!$existingHub) {
            Register-PnPHubSite -Site $newSiteFullUrl
            Set-PnPHubSite -Identity $newSiteFullUrl -Title $hubName
            logOutput "   Registered $provShortcode as hub."
        }
        else {
            logOutput "   $provShortcode already registered as hub."
        }
    }
    catch {
        logOutput "   Error registering $provShortcode as hub. Error: $_"
    }
}

function AssociateToHub ($newSiteFullUrl, $hubName, $hubAssoc) {
    logOutput "   Associating $provShortcode to hub."
    try {
        if ($hubName -and $hubAssoc) {
            $parentHub = Get-PnPHubSite | Where-Object {$_.Title -eq $hubAssoc}
            Add-PnPHubToHubAssociation -SourceUrl $newSiteFullUrl -TargetUrl $parentHub.SiteUrl
        } 
        elseif ($hubAssoc) {
            $parentHub = Get-PnPHubSite | Where-Object {$_.Title -eq $hubAssoc}
            Add-PnPHubSiteAssociation -Site $newSiteFullUrl -HubSite $parentHub.SiteUrl
        }
        logOutput "   Associated $provShortcode to hub."
    }
    catch {
        logOutput "   Error associating $provShortcode to hub. Error: $_"
    }
}

function ShareGateCopyStructure ($provSourceSite, $newSiteFullUrl, $provShortcode) {
    try {
        logOutput "   Copying list structure on $provShortcode..."
        $listsRequired = $listsCsv | Where-Object {$_.SourceSite -eq $provSourceSite -and $_.DestinationSite -eq $provShortcode}
        if ($listsRequired) {
            $srcSite = Connect-Site -Url $provSourceSite -UserName $sgUsername -Password  $secureStringPassword 
            $dstSite = Connect-Site -Url $newSiteFullUrl -UseCredentialsFrom $sgConnection
            $provSourceSiteWinAuth = $provSourceSite.Insert($winauthIndex, "-winauth")
            $pnpConnection = Connect-PnPOnline -Url $provSourceSiteWinAuth -Credentials $cred
            foreach($list in $listsRequired) {
                $provListName = $list.LibraryName
                try {
                    logOutput "      Copying list $provListName on $provShortcode..."
                    $listStartTime = Get-Date
                    $results = Copy-List -Name $provListName -SourceSite $srcSite -DestinationSite $dstSite -TaskName "$provSourceSite - $newSiteFullUrl : $provListName" -NoContent  -NoSiteFeatures -NoWorkflows
                    $listFinishTime = Get-Date
                    $listDuration = New-TimeSpan -Start $listStartTime -End $listFinishTime
                    $listDurationMinutes = $listDuration.Minutes
                    <#$srcList = Get-PnPList -Identity $provListName -Connection $pnpConnection
                    if($srcList.OnQuickLaunch) {
                        Connect-PnPOnline -Url $newSiteFullUrl -Interactive
                        $newList = Get-PnPList -Identity $provListName
                        $Context = Get-PnPContext
                        $newList.OnQuickLaunch = $True
                        $newList.Update() 
                        $Context.ExecuteQuery()
                    }#>
                    logOutput "      Copied list $provListName on $provShortcode..."
                    # Write Summary to CSV File
                    $script:listCsvData += [PSCustomObject]@{
                        SourceSite = $provSourceSite
                        DestinationSite = $newSiteFullUrl
                        ListName = $provListName
                        StartTime = $listStartTime
                        FinishTime = $listFinishTime
                        TotalProvisioningTimeMinutes = $listDurationMinutes
                        Result = $results.Result
                        Successes = $results.Successes
                        Warnings = $results.Warnings
                        Errors = $results.Errors
                    }
                }
                catch {
                    logOutput "      Error copying list $provListName on $provShortcode. Error: $_"
                }
            }
        }
    }
    catch {
        logOutput "   Error copying list structure on $provShortcode. Error: $_"
    }
}

function PostSharegateConfiguration () {
    try {
        logOutput "   Applying post ShareGate configuration to $provShortcode..."
        # Additional custom fields
        try {
            logOutput "      Adding legacy capture fields to $provShortcode..."
            foreach($field in $customFields) {
                $fieldFound = Get-PnPField -Identity $field.InternalName -ErrorAction SilentlyContinue
                $fieldName = $field.DisplayName
                if(!$fieldFound) {
                    Add-PnPField -Type $field.Type -InternalName $field.InternalName -DisplayName $field.DisplayName -Group $field.Group
                    logOutput "         Added field $fieldName to $provShortcode."
                }
                else {
                    logOutput "         Field $fieldName already exists on $provShortcode."
                }
            }
            logOutput "      Finished adding legacy capture fields to $provShortcode."
        }
        catch {
            logOutput "      Error adding legacy capture fields to $provShortcode. Error: $_"
            logOutput "   Error completing post ShareGate configuration on $provShortcode. Error: $_"
        }
        # Versioning settings & fields
        try {
            logOutput "      Setting versioning settings on $provShortcode..."
            $listsLibraries = Get-PnPList | Where-Object {$_.Hidden -eq $false -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library" -and $_.Title -ne "Site Pages"}
            foreach($list in $listsLibraries) {
                if ($list.BaseType -eq "DocumentLibrary") {
                    Set-PnPList -Identity $list -MajorVersions 500 -MinorVersions 2 -EnableVersioning $True -EnableMinorVersions $True
                } elseif ($list.BaseType -eq "GenericList") {
                    Set-PnPList -Identity $list -MajorVersions 500 -EnableVersioning $True
                }
            }
        }
        catch {
            logOutput "      Error setting versioning settings on $provShortcode. Error: $_"
            logOutput "   Error completing post ShareGate configuration on $provShortcode. Error: $_"
        }
    }
    catch {
        logOutput "   Error completing post ShareGate configuration on $provShortcode. Error: $_"
    }
}

function ProvisionSiteFromTemplate ($provSiteName, $provShortcode, $provDesc, $provHub, $provHubAssoc, $provTemplate, $provSourceSiteUrl) {
    logOutput "# Starting provisioning for site $provShortcode..."
    $siteStartTime = Get-Date

    # Request template type
    switch ($provTemplate) {
        "Hub Template" {
            $baseTemplate = "CommunicationSite"
            $provSiteFullUrl = $baseURL + "/Sites/" + $provShortcode
            $provTemplatePath = $templateFolderPath + "\pilotHub.xml"
            $teamsEnabled = $false
        }
        "Community Site Template" {
            $baseTemplate = "TeamSite"
            $provSiteFullUrl = $baseURL + "/teams/" + $provShortcode
            $provTemplatePath = $templateFolderPath + "\pilotCommunity.xml"
            $teamsEnabled = $true
        }
        "Governance Site Template" {
            $baseTemplate = "TeamSite"
            $provSiteFullUrl = $baseURL + "/Sites/" + $provShortcode
            $provTemplatePath = $templateFolderPath + "\pilotGovernance.xml"
            $teamsEnabled = $false
        }
        "Project Site Template" {
            $baseTemplate = "TeamSite"
            $provSiteFullUrl = $baseURL + "/Sites/" + $provShortcode
            $provTemplatePath = $templateFolderPath + "\pilotProject.xml"
            $teamsEnabled = $true
        }
        "Standard SharePoint Site Template" {
            $baseTemplate = "TeamSite"
            $provSiteFullUrl = $baseURL + "/Sites/" + $provShortcode
            $provTemplatePath = $templateFolderPath + "\pilotStandard.xml"
            $teamsEnabled = $false
        }
        "Standard Teams Site Template" {
            $baseTemplate = "TeamSite"
            $provSiteFullUrl = $baseURL + "/Sites/" + $provShortcode
            $provTemplatePath = $templateFolderPath + "\pilotTeams.xml"
            $teamsEnabled = $true
        }
    }

    # Provision site
    #ProvisionSite $baseTemplate $provSiteName $provShortcode $provSiteFullUrl

    # Site configuration
    #SiteConfiguration $provShortcode $provSiteFullUrl $provTemplate $provTemplatePath

    # Teamify site
    if ($teamsEnabled) {
        #TeamifySite $provSiteFullUrl
    }

    if ($provHub) {
        #RegisterHub $provSiteFullUrl $provHub
    }

    #AssociateToHub $provSiteFullUrl $provHub $provHubAssoc

    # SG Copy list structure
    ShareGateCopyStructure $provSourceSiteUrl $provSiteFullUrl $provShortcode

    #PostSharegateConfiguration

    $siteFinishTime = Get-Date

    $siteDuration = New-TimeSpan -Start $siteStartTime -End $siteFinishTime

    $siteDurationMinutes = $siteDuration.Minutes

    # Write Summary to CSV File
    $script:csvData += [PSCustomObject]@{
        SiteName = $provSiteName
        ShortCode = $provShortcode
        Template = $provTemplate
        StartTime = $siteStartTime
        FinishTime = $siteFinishTime
        TotalProvisioningTimeMinutes = $siteDurationMinutes
    }

    logOutput "# Finished provisioning site $provShortcode in $siteDurationMinutes minute/s."

}

function ProvisionSites {
    # Log file
    $script:startTimestamp = Get-Date -Format "HH:mm dd-MM-yy"
    $script:startTimestampFileName = Get-Date -Format "HH-mm-dd-MM-yy"
    $script:logFilePath = $logFileFolderPath + "\" + $startTimestampFileName + "-Foundation.txt"
    $script:csvLogPath = $logFileFolderPath + "\" + $startTimestampFileName + "-Foundation.csv"
    $script:csvListLogPath = $logFileFolderPath + "\" + $startTimestampFileName + "-Foundation_Lists.csv"
    createLog

    # Read csv
    logOutput "Getting csv data..."
    try {
        ReadCsv $filePath
        logOutput "Collected csv data."
    } 
    catch {
        logOutput "Failed to get csv data."
        exit
    }

    # Prepare SG
    Import-Module Sharegate
    # Prompt the user to enter a password
    $script:sgUsername = "wlgc3\piritahidemoadmin"
    #$script:secureStringPassword = Read-Host -Prompt "Enter the password for $sgUserName" -AsSecureString
    $script:cred = Get-Credential -UserName $sgUsername -Message 'Enter password'
    $script:sgConnection = Connect-Site -Url $adminURL -Browser

    # Provision site
    foreach ($site in $csv) {
        $csvSourceSiteUrl = $site.SourceSiteUrl
        $csvSiteName = $site.SiteName
        $csvShortcode = $site.Shortcode
        $csvDesc = $site.Description
        $csvHub = $site.HubName
        $csvHubAssoc = $site.HubAssociation
        $csvTemplate = $site.Template
        ProvisionSiteFromTemplate $csvSiteName $csvShortcode $csvDesc $csvHub $csvHubAssoc $csvTemplate $csvSourceSiteUrl
    }

    $csvData | Export-Csv -Path $csvLogPath
    $listCsvData | Export-Csv -Path $csvListLogPath

    logOutput "### Script finished ###"
}

ProvisionSites