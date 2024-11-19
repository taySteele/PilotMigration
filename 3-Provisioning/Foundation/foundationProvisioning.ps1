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
    $script:csv = Import-Csv -Path "C:\Users\tsteele\OneDrive - Capgemini\Desktop\MPI\FoundationBuildProvisioning.csv"
}

function ProvisionSite ($provBaseTemplate, $newSiteName, $newSiteShortCode, $newSiteFullUrl) {
    logOutput "Provisioning site $newSiteShortCode..."

    # Connect to admin center
    try {
        Connect-PnPOnline -Url $adminURL -Interactive -ClientId $clientId
    }
    catch {
        logOutput "Unable to connect to Admin Center"
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
            
            logOutput "Finished provisioning site $newSiteShortCode."
        } 
        else {
            $script:newSite = $existingSite
            logOutput "Site already exists for $newSiteShortCode"
        }
    } catch {
        logOutput "Failed to provision site $newSiteShortCode. Error: $_"
        Exit
    }
}

function SiteConfiguration ($newSiteShortCode, $newSite, $template, $templatePath) {
    logOutput "Applying configuration to site $newSiteShortCode..."
    try {
        # Connect to site
        Connect-PnPOnline -Url $newSite -Interactive -ClientId $clientId

        # Add admins
        Add-PnPSiteCollectionAdmin -Owners $owners

        # Enable document sets
        Enable-PnPFeature -Scope Site -Identity "3bae86a2-776d-499d-9db8-fa4cdc7884f8" -Force

        # Apply template - landing page and site settings
        Invoke-PnPSiteTemplate -ClearNavigation -Path $templatePath

        # Set theme
        Set-PnPWebTheme -Theme $themeName

        # Set logo
        Set-PnPSite -LogoFilePath $logoFile

        # Set permission settings
        #$siteGroups = Get-PnPSiteGroup
        #$membersGroup = $siteGroups | Where-Object {$_.Title -like "*Members"}
        #Set-PnPGroupPermissions -Identity $membersGroup.Title -AddRole Contribute -RemoveRole Edit

        # Set sharing settings
        $web = Get-PnPWeb -Includes MembersCanShare, AssociatedMemberGroup.AllowMembersEditMembership
        $web.MembersCanShare=$true
        $web.AssociatedMemberGroup.AllowMembersEditMembership=$false
        $web.AssociatedMemberGroup.Update()
        $web.RequestAccessEmail = $null
        $web.Update()
        $web.Context.ExecuteQuery()

        # Disable Workflow Task Content Type
        Disable-PnPFeature -Identity 57311b7a-9afd-4ff0-866e-9393ad6647b1 -Force

        # Disable Three-state workflow
        Disable-PnPFeature -Scope Site -Identity fde5d850-671e-4143-950a-87b473922dc7 -Force

        # Configure Doc ID prefix
        Set-PnPTenantSite -Identity $newSite -DenyAddAndCustomizePages:$false
        Set-PnPPropertyBagValue -Key "docid_msft_hier_siteprefix" -Value "MPIDOCID"
        Set-PnPTenantSite -Identity $newSite -DenyAddAndCustomizePages:$true

        # Remove recent link from nav
        $navigationNodes = Get-PnPNavigationNode
        if ($navigationNodes | Where-Object {$_.Title -eq "Recent"}) {
            $recentLink = $navigationNodes | Where-Object {$_.Title -eq "Recent"}
            Remove-PnPNavigationNode -Identity $recentLink -Force
        }
    
        logOutput "Finished applying configuration to site $newSiteShortCode."
    } catch {
        logOutput "Failed to apply configuration to site $newSiteShortCode. Error $_"
        Exit
    }
}

function TeamifySite ($newSite) {
    try {
        logOutput "Associating Microsoft Team to site"
        $site = Get-PnPTenantSite -Identity $newSite
        New-PnPTeamsTeam -GroupId $site.GroupId.Guid
        logOutput "Microsoft Team associated to site"
    } catch {
        logOutput "Failed to connect site to Microsoft Team. Error $_"
        Exit
    }
    
}

function RegisterHub ($newSiteFullUrl, $hubName) {
    logOutput "Registering $provShortcode as hub."
    try {
        $existingHub = Get-PnPHubSite | Where-Object {$_.Title -eq $hubName}
        if (!$existingHub) {
            Register-PnPHubSite -Site $newSiteFullUrl
            Set-PnPHubSite -Identity $newSiteFullUrl -Title $hubName
            logOutput "Registered $provShortcode as hub."
        }
        else {
            logOutput "$provShortcode already registered as hub."
        }
    }
    catch {
        logOutput "Error registering $provShortcode as hub. Error: $_"
    }
}

function AssociateToHub ($newSiteFullUrl, $hubName, $hubAssoc) {
    logOutput "Associating $provShortcode to hub."
    try {
        if ($hubName -and $hubAssoc) {
            $parentHub = Get-PnPHubSite | Where-Object {$_.Title -eq $hubAssoc}
            Add-PnPHubToHubAssociation -SourceUrl $newSiteFullUrl -TargetUrl $parentHub.SiteUrl
        } 
        elseif ($hubAssoc) {
            $parentHub = Get-PnPHubSite | Where-Object {$_.Title -eq $hubAssoc}
            Add-PnPHubSiteAssociation -Site $newSiteFullUrl -HubSite $parentHub.SiteUrl
        }
        logOutput "Associated $provShortcode to hub."
    }
    catch {
        logOutput "Error associating $provShortcode to hub. Error: $_"
    }
}

function ProvisionSiteFromTemplate ($provSiteName, $provShortcode, $provDesc, $provHub, $provHubAssoc, $provTemplate) {
    logOutput "Staring provisioning site $provShortcode."

    # Request template type
    switch ($provTemplate) {
        "Hub Template" {
            $baseTemplate = "CommunicationSite"
            $provSiteFullUrl = $baseURL + "/Sites/" + $provShortcode
            $provTemplatePath = $templateFolderPath + "\foundationHub.xml"
        }
        "Community Site Template" {
            $baseTemplate = "TeamSite"
            $provSiteFullUrl = $baseURL + "/teams/" + $provShortcode
            $provTemplatePath = $templateFolderPath + "\foundationCommunity.xml"
        }
    }

    # Provision site
    ProvisionSite $baseTemplate $provSiteName $provShortcode $provSiteFullUrl

    # Site configuration
    #SiteConfiguration $provShortcode $provSiteFullUrl $provTemplate $provTemplatePath

    # Teamify site
    if ($provTemplate -eq "Community Site Template") {
        #TeamifySite $provSiteFullUrl
    }

    if ($provHub) {
        RegisterHub $provSiteFullUrl $provHub
    }

    AssociateToHub $provSiteFullUrl $provHub $provHubAssoc

    logOutput "Finished provisioning site $provShortcode."

}

function ProvisionSites {
    # Log file
    $script:startTimestamp = Get-Date -Format "HH:mm dd-MM-yy"
    $script:startTimestampFileName = Get-Date -Format "HH-mm-dd-MM-yy"
    $script:logFilePath = $logFileFolderPath + "\" + $startTimestampFileName + ".txt"
    createLog

    # Read csv
    logOutput "Getting csv data..."
    try {
        #$csvPath = Read-Host "Enter the path to the csv file"
        ReadCsv $csvPath
        logOutput "Collected csv data."
    } 
    catch {
        logOutput "Failed to get csv data."
        exit
    }

    # Provision site
    foreach ($site in $csv) {
        $csvSiteName = $site.SiteName
        $csvShortcode = $site.Shortcode
        $csvDesc = $site.Description
        $csvHub = $site.HubName
        $csvHubAssoc = $site.HubAssociation
        $csvTemplate = $site.Template
        ProvisionSiteFromTemplate $csvSiteName $csvShortcode $csvDesc $csvHub $csvHubAssoc $csvTemplate
    }
}

ProvisionSites