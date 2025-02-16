# Import required modules and assemblies
Import-Module ActiveDirectory

# Check PowerShell version and load assemblies accordingly
$PSVersion = $PSVersionTable.PSVersion.Major
if ($PSVersion -ge 6) {
    # PowerShell 7.x and above requires explicit assembly loading
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Collections
    Add-Type -AssemblyName System.Windows.Forms
} else {
    # PowerShell 5.x can use traditional assembly loading
    [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')
    [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')
    [void][System.Reflection.Assembly]::LoadWithPartialName('WindowsBase')
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
}

# Get script location - handle both running the script and dot-sourcing
$ScriptPath = $MyInvocation.MyCommand.Path
if (-not $ScriptPath) {
    $ScriptPath = $PSCommandPath
}
if (-not $ScriptPath) {
    Write-Error "Unable to determine script path"
    return
}

$ScriptDir = Split-Path -Parent $ScriptPath
Write-Host "Script Directory: $ScriptDir"

# Create log directory if it doesn't exist
$LogDir = Join-Path $ScriptDir "Logs"
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

# Initialize logging
$LogFile = Join-Path $LogDir "GroupReview_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Create runspace-safe variables
$script:LogTextBox = $null
$script:LogOverlay = $null
$script:Window = $null
$script:StopProcessing = $false

# Set up paths for resource files
$script:XamlFile = Join-Path $ScriptDir "Resources\GUI.xaml"
$script:HtmlTemplateFile = Join-Path $ScriptDir "Resources\HTML_Template.html"

Write-Host "XAML File Path: $($script:XamlFile)"
Write-Host "HTML Template Path: $($script:HtmlTemplateFile)"

# Verify resource files exist
if (-not (Test-Path $script:XamlFile)) {
    Write-Error "XAML file not found: $script:XamlFile"
    return
}

if (-not (Test-Path $script:HtmlTemplateFile)) {
    Write-Error "HTML template file not found: $script:HtmlTemplateFile"
    return
}

# Load XAML and HTML templates
try {
    Write-Host "Loading XAML from: $($script:XamlFile)"
    $xamlContent = Get-Content -Path $script:XamlFile -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($xamlContent)) {
        throw "XAML file is empty"
    }
    [xml]$script:XAML = $xamlContent
    Write-Host "XAML loaded successfully"
    Write-Log "Loaded XAML template from: $script:XamlFile"
    
    Write-Host "Loading HTML template from: $($script:HtmlTemplateFile)"
    $script:HTMLTemplate = Get-Content -Path $script:HtmlTemplateFile -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($script:HTMLTemplate)) {
        throw "HTML template file is empty"
    }
    Write-Host "HTML template loaded successfully"
    Write-Log "Loaded HTML template from: $script:HtmlTemplateFile"
}
catch {
    Write-Error "Error loading resource files: $_"
    Write-Host "Error details: $($_.Exception.Message)"
    Write-Host "Stack trace: $($_.ScriptStackTrace)"
    return
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [switch]$NoConsole,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Type = 'Info'
    )
    try {
        $LogMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
        Add-Content -Path $LogFile -Value $LogMessage -ErrorAction Stop
        
        if (-not $NoConsole) {
            switch ($Type) {
                'Error'   { Write-Host $LogMessage -ForegroundColor Red }
                'Warning' { Write-Host $LogMessage -ForegroundColor Yellow }
                'Success' { Write-Host $LogMessage -ForegroundColor Green }
                default   { Write-Host $LogMessage }
            }
        }
        
        if ($script:LogTextBox -and $script:Window) {
            $script:Window.Dispatcher.Invoke(
                [Action]{
                    # Insert new text at the beginning with appropriate color
                    $newText = New-Object System.Windows.Documents.Run
                    $newText.Text = "$LogMessage`n"
                    
                    # Set text color based on message type
                    $newText.Foreground = switch ($Type) {
                        'Error'   { [System.Windows.Media.Brushes]::Red }
                        'Warning' { [System.Windows.Media.Brushes]::DarkOrange }
                        'Success' { [System.Windows.Media.Brushes]::Green }
                        default   { [System.Windows.Media.Brushes]::Black }
                    }
                    
                    $script:LogTextBox.Text = "$LogMessage`n$($script:LogTextBox.Text)"
                    
                    # Trim to last 250 entries if needed
                    $currentLines = $script:LogTextBox.Text -split "`n"
                    if ($currentLines.Count -gt 250) {
                        $script:LogTextBox.Text = ($currentLines | Select-Object -First 250) -join "`n"
                    }
                    
                    $script:LogTextBox.ScrollToHome()
                },
                [System.Windows.Threading.DispatcherPriority]::Background
            )
        }
    }
    catch {
        Write-Host "Error in Write-Log: $_" -ForegroundColor Red
    }
}

# Function to get the downloads folder path
function Get-DownloadsFolder {
    try {
        $shell = New-Object -ComObject Shell.Application
        $downloads = $shell.NameSpace('shell:Downloads').Self.Path
        if ($downloads -and (Test-Path $downloads)) {
            Write-Log "Using downloads folder: $downloads"
            return $downloads
        }
        Write-Log "Downloads folder not found, using script directory"
        return $ScriptDir
    }
    catch {
        Write-Log "Error finding downloads folder: $_"
        return $ScriptDir
    }
}

# Function to analyze group health
function Get-GroupHealth {
    param($Group)
    
    $health = @{
        Issues = New-Object System.Collections.ArrayList
        Score = 100  # Start with perfect score
    }
    
    # Check for empty description
    if ([string]::IsNullOrWhiteSpace($Group.Description)) {
        [void]$health.Issues.Add("Missing description")
        $health.Score = $health.Score - 20
    }
    
    # Check for missing manager
    if ([string]::IsNullOrWhiteSpace($Group.Manager)) {
        [void]$health.Issues.Add("No manager assigned")
        $health.Score = $health.Score - 20
    }
    
    # Check member count
    if ($Group.TotalMembers -eq 0) {
        [void]$health.Issues.Add("Empty group")
        $health.Score = $health.Score - 30
    }
    elseif ($Group.TotalMembers -gt 1000) {
        [void]$health.Issues.Add("Large group (>1000 members)")
        $health.Score = $health.Score - 10
    }
    
    # Check age
    $age = (Get-Date) - $Group.Created
    if ($age.Days -gt 365 * 2) {
        [void]$health.Issues.Add("Group older than 2 years")
        $health.Score = $health.Score - 10
    }
    
    # Check disabled user percentage
    if ($Group.UserMembers -gt 0) {
        $disabledPercentage = ($Group.DisabledMembers / $Group.UserMembers) * 100
        if ($disabledPercentage -gt 40) {
            [void]$health.Issues.Add("High percentage of disabled users (>40%)")
            $health.Score = $health.Score - 30
        }
        elseif ($disabledPercentage -gt 20) {
            [void]$health.Issues.Add("Moderate percentage of disabled users (>20%)")
            $health.Score = $health.Score - 15
        }
    }
    
    # Ensure score doesn't go below 0
    $health.Score = [Math]::Max(0, $health.Score)
    
    return $health
}

# Function to get available OUs
function Get-ADOUList {
    Write-Log "Getting list of OUs..."
    try {
        $domain = Get-ADDomain
        $ous = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName -SearchBase $domain.DistinguishedName |
            ForEach-Object {
                # Get group count for this OU
                $groupCount = @(Get-ADGroup -Filter * -SearchBase $_.DistinguishedName -SearchScope OneLevel).Count
                
                # Format the OU path for display - replace DC parts and format OU path with arrows
                $ouPath = $_.DistinguishedName
                $ouPath = $ouPath -replace '(,DC=[\w-]+)+$', ''  # Remove DC components
                $ouPath = $ouPath -replace ',OU=', ' -> '         # Replace OU separators with arrows
                $ouPath = $ouPath -replace '^OU=', ''            # Remove leading OU=
                
                [PSCustomObject]@{
                    Name = "$ouPath ($groupCount groups)"
                    DistinguishedName = $_.DistinguishedName
                    Description = "Full Path: $($_.DistinguishedName)"
                    GroupCount = $groupCount
                }
            } | Sort-Object { $_['GroupCount'] } -Descending
        
        return $ous
    } 
    catch {
        Write-Log "Error getting OU list: $_"
        return @()
    }
}

# Function to get nested group membership without cycles
function Get-NestedGroupMembership {
    param(
        [string]$GroupDN,
        [System.Collections.Generic.HashSet[string]]$ProcessedGroups = $null,
        [System.Collections.Generic.HashSet[string]]$AllNestedGroups = $null
    )
    
    if ($null -eq $ProcessedGroups) {
        $ProcessedGroups = New-Object System.Collections.Generic.HashSet[string]
    }
    if ($null -eq $AllNestedGroups) {
        $AllNestedGroups = New-Object System.Collections.Generic.HashSet[string]
    }
    
    # If we've already processed this group, skip it to avoid cycles
    if (-not $ProcessedGroups.Add($GroupDN)) {
        return $AllNestedGroups
    }
    
    try {
        $group = Get-ADGroup -Identity $GroupDN -Properties memberOf
        
        if ($null -ne $group.memberOf) {
            # Convert memberOf to array safely based on its type
            $memberOfGroups = if ($group.memberOf -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                [array]($group.memberOf)
            } elseif ($group.memberOf -is [string]) {
                @($group.memberOf)
            } else {
                [array]($group.memberOf)
            }
            
            foreach ($memberOfGroup in $memberOfGroups) {
                if ($AllNestedGroups.Add($memberOfGroup)) {
                    # Only recurse if this is a new group
                    Get-NestedGroupMembership -GroupDN $memberOfGroup -ProcessedGroups $ProcessedGroups -AllNestedGroups $AllNestedGroups | Out-Null
                }
            }
        }
        
        return $AllNestedGroups
    }
    catch {
        Write-Log "Error getting nested groups for ${GroupDN}: $_" -NoConsole
        return $AllNestedGroups
    }
}

# Update the group details collection to include nested group information
function Get-GroupDetails {
    param(
        [string[]]$SelectedOUs
    )
    Write-Log "Retrieving AD groups for selected OUs..."
    
    try {
        $domain = Get-ADDomain
        Write-Log "Connected to domain: $($domain.DNSRoot)"
        
        $allGroups = New-Object System.Collections.ArrayList
        $processedGroups = 0
        [int]$totalGroups = 0
        
        # Create hashtable to store OU statistics
        ${script:OUStats} = @{}
        
        # First pass - count total groups
        foreach($ou in $SelectedOUs) {
            # Use ErrorAction to handle PS7 error behavior
            $groups = @(Get-ADGroup -Filter * -SearchBase $ou -ErrorAction Stop)
            $totalGroups = $totalGroups + $groups.Count
            
            # Initialize OU stats
            ${script:OUStats}[$ou] = @{
                GroupCount = $groups.Count
                EnabledMembers = 0
                DisabledMembers = 0
                TotalMembers = 0
                DisabledPercentage = 0
                NestedGroupCount = 0
                MaxNestingDepth = 0
            }
        }
        
        Write-Log "Total groups to process: $totalGroups"
        
        foreach($ou in $SelectedOUs) {
            Write-Log "Processing OU: $ou"
            $groups = Get-ADGroup -Filter * -SearchBase $ou -Properties Description, Info, whenCreated, 
                managedBy, mail, groupCategory, groupScope, member, memberOf, 
                DistinguishedName, objectSid, sAMAccountName
            
            Write-Log "Found $($groups.Count) groups in $ou"
            
            foreach($group in $groups) {
                if ($script:StopProcessing) {
                    Write-Log "Processing cancelled by user"
                    return $null
                }
                
                $processedGroups++
                $percentComplete = [math]::Round(($processedGroups / $totalGroups) * 100, 1)
                Write-Log "Processing group ($processedGroups/$totalGroups - $percentComplete%): $($group.Name)" -NoConsole
                
                try {
                    # Get nested group information
                    $nestedGroups = Get-NestedGroupMembership -GroupDN $group.DistinguishedName
                    $nestingDepth = if ($nestedGroups -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                        ([array]$nestedGroups).Count
                    } else {
                        if ($null -eq $nestedGroups) { 0 } else { @($nestedGroups).Count }
                    }
                    
                    # Update OU statistics for nested groups - ensure numeric operations
                    $currentNestedCount = [int](${script:OUStats}[$ou].NestedGroupCount)
                    ${script:OUStats}[$ou].NestedGroupCount = $currentNestedCount + $nestingDepth
                    ${script:OUStats}[$ou].MaxNestingDepth = [Math]::Max([int](${script:OUStats}[$ou].MaxNestingDepth), $nestingDepth)
                    
                    # Initialize member counts
                    $userMembers = 0
                    $groupMembers = 0
                    $computerMembers = 0
                    $totalMembers = 0
                    $enabledMembers = 0
                    $disabledMembers = 0
                    
                    if ($group.member) {
                        Write-Log "Getting members for group: $($group.Name)" -NoConsole
                        
                        # Get all user members and their enabled status
                        $users = Get-ADUser -LDAPFilter "(memberOf=$($group.DistinguishedName))" -Properties Enabled -ResultSetSize $null
                        $userMembers = if ($users -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                            ([array]$users).Count
                        } else {
                            if ($null -eq $users) { 0 } else { @($users).Count }
                        }
                        
                        $enabledMembers = if ($users) {
                            @($users | Where-Object { $_.Enabled }).Count
                        } else { 0 }
                        
                        $disabledMembers = if ($users) {
                            @($users | Where-Object { -not $_.Enabled }).Count
                        } else { 0 }
                        
                        Start-Sleep -Milliseconds 50
                        
                        $groupMembers = @(Get-ADGroup -LDAPFilter "(memberOf=$($group.DistinguishedName))" -ResultSetSize $null).Count
                        Start-Sleep -Milliseconds 50
                        
                        $computerMembers = @(Get-ADComputer -LDAPFilter "(memberOf=$($group.DistinguishedName))" -ResultSetSize $null).Count
                        Start-Sleep -Milliseconds 50
                        
                        $totalMembers = $userMembers + $groupMembers + $computerMembers
                        
                        # Update OU statistics - ensure numeric operations
                        $currentEnabled = [int](${script:OUStats}[$ou].EnabledMembers)
                        $currentDisabled = [int](${script:OUStats}[$ou].DisabledMembers)
                        $currentTotal = [int](${script:OUStats}[$ou].TotalMembers)
                        
                        ${script:OUStats}[$ou].EnabledMembers = $currentEnabled + $enabledMembers
                        ${script:OUStats}[$ou].DisabledMembers = $currentDisabled + $disabledMembers
                        ${script:OUStats}[$ou].TotalMembers = $currentTotal + $totalMembers
                    }
                    
                    # Convert manager CN to UPN if present
                    $managerUPN = if ($group.managedBy) {
                        try {
                            Write-Log "Getting manager info for group: $($group.Name)" -NoConsole
                            $managerUser = Get-ADUser -Identity $group.managedBy -Properties UserPrincipalName, DisplayName, Title
                            [PSCustomObject]@{
                                UPN = $managerUser.UserPrincipalName
                                DisplayName = $managerUser.DisplayName
                                Title = $managerUser.Title
                            }
                        } catch {
                            Write-Log "Error getting manager for group $($group.Name): $_" -NoConsole
                            $group.managedBy
                        }
                    } else { $null }
                    
                    # Create group object
                    $groupObj = [PSCustomObject]@{
                        Name = $group.Name
                        Description = $group.Description
                        Info = $group.Info
                        TotalMembers = $totalMembers
                        UserMembers = $userMembers
                        GroupMembers = $groupMembers
                        ComputerMembers = $computerMembers
                        Created = $group.whenCreated
                        Manager = $managerUPN
                        Email = $group.mail
                        Category = $group.groupCategory
                        Scope = $group.groupScope
                        NestedInGroupCount = if ($null -ne $group.memberOf) {
                            if ($group.memberOf -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                                [array]($group.memberOf).Count
                            } else {
                                if ($group.memberOf -is [string]) { 1 } else { @($group.memberOf).Count }
                            }
                        } else { 0 }
                        HasNestedGroups = ($groupMembers -gt 0)
                        DN = $group.DistinguishedName
                        OU = ($group.DistinguishedName -split ',',2)[1]
                        SamAccountName = $group.sAMAccountName
                        # Add parent group names - handle ADPropertyValueCollection properly
                        ParentGroups = [array](Get-ADGroup -LDAPFilter "(member=$($group.DistinguishedName))" -Properties name | 
                            Select-Object -ExpandProperty name | Where-Object { $_ })
                        # Add nested group names - handle ADPropertyValueCollection properly
                        NestedGroups = [array](Get-ADGroup -LDAPFilter "(memberOf=$($group.DistinguishedName))" -Properties name | 
                            Select-Object -ExpandProperty name | Where-Object { $_ })
                    }
                    
                    # Add health check
                    $health = Get-GroupHealth $groupObj
                    $groupObj | Add-Member -NotePropertyName HealthScore -NotePropertyValue $health.Score
                    $groupObj | Add-Member -NotePropertyName HealthIssues -NotePropertyValue $health.Issues
                    
                    [void]$allGroups.Add($groupObj)
                }
                catch {
                    Write-Log "Error processing group $($group.Name): $_" -Type Error
                    continue
                }
            }
        }
        
        Write-Log "Processed all groups successfully"
        return $allGroups
    }
    catch {
        Write-Log "Error retrieving group details: $_"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
        return $null
    }
}

# Function to generate HTML report
function New-HTMLReport {
    param($Groups)
    
    try {
        $DownloadsFolder = Get-DownloadsFolder
        $TimeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $ReportFile = Join-Path $DownloadsFolder "ADGroupReview_$TimeStamp.html"
        $CSVFile = Join-Path $DownloadsFolder "ADGroupReview_$TimeStamp.csv"
        Write-Log "Generating reports: HTML and CSV"
        
        # Export CSV with formatted data
        Write-Log "Generating CSV export..."
        $Groups | ForEach-Object {
            [PSCustomObject]@{
                'Group Name' = $_.Name
                'SAM Account Name' = $_.SamAccountName
                'Description' = $_.Description
                'Health Score' = $_.HealthScore
                'Health Issues' = ($_.HealthIssues -join '; ')
                'Total Members' = $_.TotalMembers
                'User Members' = $_.UserMembers
                'Group Members' = $_.GroupMembers
                'Computer Members' = $_.ComputerMembers
                'Manager Name' = $_.Manager.DisplayName
                'Manager Title' = $_.Manager.Title
                'Manager Email' = $_.Manager.UPN
                'Created Date' = $_.Created.ToString('yyyy-MM-dd')
                'Category' = $_.Category
                'Scope' = $_.Scope
                'Nested In Groups' = $_.NestedInGroupCount
                'Has Nested Groups' = $_.HasNestedGroups
                'Email' = $_.Email
                'Organizational Unit' = $_.OU
                'Distinguished Name' = $_.DN
                'Notes' = $_.Info
            }
        } | Export-Csv -Path $CSVFile -NoTypeInformation -Encoding UTF8
        Write-Log "CSV export saved to: $CSVFile"

        # Calculate statistics
        $totalGroups = $Groups.Count
        $emptyGroups = @($Groups | Where-Object { $_.TotalMembers -eq 0 } | Sort-Object -Unique DN).Count
        $noManager = @($Groups | Where-Object { -not $_.Manager } | Sort-Object -Unique DN).Count
        $noDescription = @($Groups | Where-Object { -not $_.Description } | Sort-Object -Unique DN).Count
        $nestedGroups = @($Groups | Where-Object { $_.HasNestedGroups } | Sort-Object -Unique DN).Count
        $avgHealth = ($Groups | Sort-Object -Unique DN | Measure-Object -Property HealthScore -Average).Average
        $criticalGroups = @($Groups | Where-Object { $_.HealthScore -le 50 } | Sort-Object -Unique DN).Count
        $warningGroups = @($Groups | Where-Object { $_.HealthScore -gt 50 -and $_.HealthScore -le 80 } | Sort-Object -Unique DN).Count
        $healthyGroups = @($Groups | Where-Object { $_.HealthScore -gt 80 } | Sort-Object -Unique DN).Count
        
        # Calculate additional statistics
        $totalMembers = ($Groups | Measure-Object -Property TotalMembers -Sum).Sum
        $activeMembers = ($Groups | Measure-Object -Property UserMembers -Sum).Sum
        $disabledMembers = ($Groups | Where-Object { -not $_.Manager } | Measure-Object -Property UserMembers -Sum).Sum
        $oldestGroup = $Groups | Sort-Object Created | Select-Object -First 1
        $newestGroup = $Groups | Sort-Object Created -Descending | Select-Object -First 1

        # Process OU statistics
        $ouStats = ${script:OUStats}.GetEnumerator() | ForEach-Object {
            $fullDN = $_.Key
            $stats = $_.Value
            
            # Split DN and get OU parts
            $parts = $fullDN -split ',' | Where-Object { $_ -match '^(OU|DC)=' }
            $ouParts = @($parts | Where-Object { $_ -match '^OU=' })
            $currentOU = ($ouParts[0] -replace '^OU=','').Trim()
            $parentOU = if ($ouParts.Count -gt 1) {
                ($ouParts[1] -replace '^OU=','').Trim()
            } else { 
                $null 
            }
            
            # Only include if this is a child OU (has no child OUs under it)
            $isChildOU = -not (${script:OUStats}.Keys | Where-Object { 
                $_ -ne $fullDN -and $_ -like "*,$fullDN"
            })
            
            if ($isChildOU) {
                # Calculate disabled percentage
                $disabledPercentage = if ($stats.TotalMembers -gt 0) {
                    [math]::Round(($stats.DisabledMembers / $stats.TotalMembers) * 100, 1)
                } else { 
                    0 
                }
                
                @{
                    CurrentOU = $currentOU
                    ParentOU = $parentOU
                    FullDN = $fullDN
                    GroupCount = $stats.GroupCount
                    EnabledMembers = $stats.EnabledMembers
                    DisabledMembers = $stats.DisabledMembers
                    TotalMembers = ($stats.EnabledMembers + $stats.DisabledMembers)
                    DisabledPercentage = $disabledPercentage
                    NestedGroupCount = $stats.NestedGroupCount
                    MaxNestingDepth = $stats.MaxNestingDepth
                }
            }
        } | Where-Object { $_ -ne $null } | Sort-Object { $_.GroupCount } -Descending

        # Calculate totals using only child OU data
        $totalMembers = ($ouStats | Measure-Object -Property TotalMembers -Sum).Sum
        $activeMembers = ($ouStats | Measure-Object -Property EnabledMembers -Sum).Sum
        $disabledMembers = ($ouStats | Measure-Object -Property DisabledMembers -Sum).Sum

        # Format groups for template
        $formattedGroups = $Groups | Sort-Object -Unique Name | Sort-Object @{Expression={$_.TotalMembers}; Descending=$true}, Name | ForEach-Object {
            $healthClass = if ($_.HealthScore -le 50) {
                'badge-critical'
            } elseif ($_.HealthScore -le 80) {
                'badge-warning'
            } else {
                'badge-success'
            }

            @{
                ID = [System.Web.HttpUtility]::HtmlEncode($_.DN)
                Name = [System.Web.HttpUtility]::HtmlEncode($_.Name)
                SamAccountName = [System.Web.HttpUtility]::HtmlEncode($_.SamAccountName)
                Description = [System.Web.HttpUtility]::HtmlEncode($_.Description)
                HealthScore = $_.HealthScore
                HealthClass = $healthClass
                TotalMembers = $_.TotalMembers
                UserMembers = $_.UserMembers
                GroupMembers = $_.GroupMembers
                ComputerMembers = $_.ComputerMembers
                EnabledUsers = $_.EnabledUsers
                DisabledUsers = $_.DisabledUsers
                Manager = if ($_.Manager) {
                    @{
                        DisplayName = [System.Web.HttpUtility]::HtmlEncode($_.Manager.DisplayName)
                        Title = [System.Web.HttpUtility]::HtmlEncode($_.Manager.Title)
                        UPN = [System.Web.HttpUtility]::HtmlEncode($_.Manager.UPN)
                    }
                } else { 
                    $null 
                }
                Created = $_.Created.ToString('yyyy-MM-dd')
                Category = [System.Web.HttpUtility]::HtmlEncode($_.Category)
                Scope = [System.Web.HttpUtility]::HtmlEncode($_.Scope)
                Email = [System.Web.HttpUtility]::HtmlEncode($_.Email)
                DN = [System.Web.HttpUtility]::HtmlEncode($_.DN)
                OU = [System.Web.HttpUtility]::HtmlEncode($_.OU)
                HealthIssues = @($_.HealthIssues | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) })
                ParentGroups = @($_.ParentGroups | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) })
                NestedGroups = @($_.NestedGroups | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) })
                NestingDepth = $_.NestingDepth
                NestingWarning = $_.NestingDepth -gt 5
                HasNestedGroups = $_.HasNestedGroups
            }
        }

        # Calculate percentages
        $noManagerPercent = [Math]::Round(($noManager / $totalGroups) * 100, 1)
        $noDescriptionPercent = [Math]::Round(($noDescription / $totalGroups) * 100, 1)

        # Create template data object
        $templateData = @{
            # Report Metadata
            REPORT_DATE = (Get-Date -Format "MMMM d, yyyy  •  HH:mm")
            
            # Group Overview
            TOTAL_GROUPS = $totalGroups
            EMPTY_GROUPS = $emptyGroups
            AVG_GROUPS_PER_OU = [Math]::Round(($totalGroups / ($ouStats.Count + 0.0)), 1)
            
            # Health Status
            AVG_HEALTH = [Math]::Round($avgHealth, 1)
            CRITICAL_GROUPS = $criticalGroups
            WARNING_GROUPS = $warningGroups
            HEALTHY_GROUPS = $healthyGroups
            
            # OU Statistics
            TOTAL_OUS = $ouStats.Count
            EMPTY_OUS = @($ouStats | Where-Object { $_.GroupCount -eq 0 }).Count
            MAX_GROUPS_OU = ($ouStats | Measure-Object -Property GroupCount -Maximum).Maximum
            
            # User Distribution
            TOTAL_MEMBERS = $totalMembers
            DISABLED_USERS = $disabledMembers
            ACTIVE_MEMBERS = $activeMembers
            USER_DISTRIBUTION = @{
                Enabled = $activeMembers
                Disabled = $disabledMembers
            }
            
            # Management Status
            NO_MANAGER = $noManager
            NO_DESCRIPTION = $noDescription
            NO_MANAGER_PERCENT = $noManagerPercent
            NO_DESCRIPTION_PERCENT = $noDescriptionPercent
            
            # Group Structure
            NESTED_GROUPS = $nestedGroups
            MAX_NESTING_DEPTH = ($Groups | Measure-Object -Property NestingDepth -Maximum).Maximum
            GROUP_CATEGORIES = "Security: $totalGroups"
            SCOPE_DISTRIBUTION = "Global: $totalGroups"
            
            # Age and Size Analysis
            OLDEST_GROUP = @{
                Name = ($Groups | Sort-Object Created | Select-Object -First 1).Name
                Created = ($Groups | Sort-Object Created | Select-Object -First 1).Created.ToString('yyyy-MM-dd')
            }
            LARGEST_GROUP = @{
                Name = ($Groups | Sort-Object TotalMembers -Descending | Select-Object -First 1).Name
                Members = ($Groups | Sort-Object TotalMembers -Descending | Select-Object -First 1).TotalMembers
            }
            LARGEST_OU = @{
                Name = ($ouStats | Sort-Object GroupCount -Descending | Select-Object -First 1).CurrentOU
                Groups = ($ouStats | Sort-Object GroupCount -Descending | Select-Object -First 1).GroupCount
            }
            
            # Full Data Sets
            GROUPS = $formattedGroups
            OU_STATS = $ouStats
        }

        # Convert template data to JSON for JavaScript
        $templateDataJson = $templateData | ConvertTo-Json -Depth 10 -Compress

        # Load and format HTML template
        $HTML = $script:HTMLTemplate
        
        # Replace template data placeholder
        $HTML = $HTML -replace 'var templateData = \{[^}]*\};', "var templateData = $templateDataJson;"
        
        # Save report with UTF8 encoding without BOM
        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
        [System.IO.File]::WriteAllLines($ReportFile, $HTML, $Utf8NoBomEncoding)
        Write-Log "Report saved to: $ReportFile"
        
        # Open the downloads folder and report in a UI-safe way
        $script:Window.Dispatcher.Invoke({
            try {
                Write-Log "Opening report location..."
                # Open folder and select both files
                $script:files = @($ReportFile, $CSVFile)
                Start-Process "explorer.exe" -ArgumentList "/select,`"$ReportFile`""
                
                Write-Log "Opening HTML report..."
                Start-Process $ReportFile
                
                Write-Log "Reports generated successfully"
                [System.Windows.MessageBox]::Show(
                    "Reports generated successfully!`n`nHTML Report: $(Split-Path $ReportFile -Leaf)`nCSV Export: $(Split-Path $CSVFile -Leaf)`n`nLocation: $DownloadsFolder", 
                    "Success",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
    }
    catch {
                Write-Log "Error opening reports: $_"
                [System.Windows.MessageBox]::Show(
                    "Reports generated but could not be opened automatically.`n`nLocation: $DownloadsFolder`n`nFiles:`n- $(Split-Path $ReportFile -Leaf)`n- $(Split-Path $CSVFile -Leaf)", 
                    "Warning"
                )
            }
        })
        
        return $true
    }
    catch {
        Write-Log "Error generating HTML report: $_"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
        $script:Window.Dispatcher.Invoke({
            [System.Windows.MessageBox]::Show("Error generating report. Check the log file for details.", "Error")
        })
        return $false
    }
}

# Function to prepare report data
function Prepare-ReportData {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Groups,
        [Parameter(Mandatory=$true)]
        [array]$OUStats
    )
    
    try {
        # Basic Group Statistics
        $totalGroups = $Groups.Count
        $emptyGroups = ($Groups | Where-Object { $_.TotalMembers -eq 0 }).Count
        
        # Health Metrics
        $avgHealth = [math]::Round(($Groups | Measure-Object -Property HealthScore -Average).Average, 1)
        $criticalGroups = ($Groups | Where-Object { $_.HealthScore -le 50 }).Count
        $warningGroups = ($Groups | Where-Object { $_.HealthScore -gt 50 -and $_.HealthScore -le 80 }).Count
        $healthyGroups = ($Groups | Where-Object { $_.HealthScore -gt 80 }).Count
        
        # OU Statistics
        $totalOUs = $OUStats.Count
        $emptyOUs = ($OUStats | Where-Object { $_.GroupCount -eq 0 }).Count
        $maxGroupsOU = ($OUStats | Measure-Object -Property GroupCount -Maximum).Maximum
        $avgGroupsPerOU = [math]::Round(($OUStats | Measure-Object -Property GroupCount -Average).Average, 1)
        
        # Member Statistics
        $totalMembers = ($Groups | Measure-Object -Property TotalMembers -Sum).Sum
        $disabledUsers = ($Groups | Measure-Object -Property DisabledUsers -Sum).Sum
        $enabledUsers = ($Groups | Measure-Object -Property EnabledUsers -Sum).Sum
        
        # Management Status
        $noManager = ($Groups | Where-Object { $null -eq $_.Manager }).Count
        $noDescription = ($Groups | Where-Object { [string]::IsNullOrEmpty($_.Description) }).Count
        $noManagerPercent = [math]::Round(($noManager / $totalGroups) * 100, 1)
        $noDescriptionPercent = [math]::Round(($noDescription / $totalGroups) * 100, 1)
        
        # Group Structure
        $nestedGroups = ($Groups | Where-Object { $_.HasNestedGroups }).Count
        $maxNestingDepth = ($Groups | Measure-Object -Property NestingDepth -Maximum).Maximum
        
        # Create the consolidated report data object
        $reportData = @{
            # Group Overview
            TOTAL_GROUPS = $totalGroups
            EMPTY_GROUPS = $emptyGroups
            AVG_GROUPS_PER_OU = $avgGroupsPerOU
            
            # Health Status
            AVG_HEALTH = $avgHealth
            CRITICAL_GROUPS = $criticalGroups
            WARNING_GROUPS = $warningGroups
            HEALTHY_GROUPS = $healthyGroups
            
            # OU Statistics
            TOTAL_OUS = $totalOUs
            EMPTY_OUS = $emptyOUs
            MAX_GROUPS_OU = $maxGroupsOU
            
            # User Distribution
            TOTAL_MEMBERS = $totalMembers
            DISABLED_USERS = $disabledUsers
            USER_DISTRIBUTION = @{
                Enabled = $enabledUsers
                Disabled = $disabledUsers
            } | ConvertTo-Json
            
            # Management Status
            NO_MANAGER = $noManager
            NO_DESCRIPTION = $noDescription
            NO_MANAGER_PERCENT = $noManagerPercent
            NO_DESCRIPTION_PERCENT = $noDescriptionPercent
            
            # Group Structure
            NESTED_GROUPS = $nestedGroups
            MAX_NESTING_DEPTH = $maxNestingDepth
            GROUP_CATEGORIES = "Security: $totalGroups"
            SCOPE_DISTRIBUTION = "Global: $totalGroups"
            
            # Full Data Sets
            GROUPS = $Groups
            OU_STATS = $OUStats
        }
        
        Write-Log -Message "Report data prepared successfully" -Type Success
        return $reportData
    }
    catch {
        Write-Log -Message "Error preparing report data: $_" -Type Error
        return $null
    }
}

# Create and show the window immediately
try {
    $Reader = [System.Xml.XmlNodeReader]::New($script:XAML)
    $Window = [Windows.Markup.XamlReader]::Load($Reader)
    
    if (-not $Window) {
        throw "Failed to create window"
    }

    # Store controls
    $script:Window = $Window
    $script:LogTextBox = $Window.FindName("logTextBox")
    $script:LogOverlay = $Window.FindName("logOverlay")
    $script:LoadingOverlay = $Window.FindName("loadingOverlay")
    $GenerateButton = $Window.FindName("btnGenerate")
    $OUList = $Window.FindName("OUList")
    $CloseLogButton = $Window.FindName("btnCloseLog")
    $ToggleSelectButton = $Window.FindName("btnToggleSelect")

    # Disable Generate button until OUs are loaded
    $GenerateButton.IsEnabled = $false
    $ToggleSelectButton.IsEnabled = $false

    # Add button click handlers
    $CloseLogButton.Add_Click({
        $script:StopProcessing = $true
        $script:LogOverlay.Visibility = "Collapsed"
    })

    $ToggleSelectButton.Add_Click({
        $isChecked = $ToggleSelectButton.IsChecked
        
        # Ensure containers are generated
        $OUList.UpdateLayout()
        Start-Sleep -Milliseconds 100
        
        # Toggle all checkboxes
        $OUList.Items | ForEach-Object {
            $container = $OUList.ItemContainerGenerator.ContainerFromItem($_)
            if ($container) {
                $checkbox = $container.ContentTemplate.FindName("ouCheckBox", $container)
                if (-not $checkbox) {
                    # If named checkbox not found, try to find first checkbox in container
                    $checkbox = [Windows.Media.VisualTreeHelper]::GetChild($container, 0) -as [System.Windows.Controls.CheckBox]
                }
                if ($checkbox) {
                    $checkbox.IsChecked = $isChecked
                }
            }
        }
    })

    $GenerateButton.Add_Click({
        Write-Log "Generate button clicked - preparing to collect selected OUs..."
        
        # Ensure containers are generated
        $OUList.UpdateLayout()
        Start-Sleep -Milliseconds 100  # Give WPF time to complete container generation
        
        try {
            # Get selected OUs - only get checked OUs
            $selectedOUs = $OUList.Items | 
                Where-Object {
                    $container = $OUList.ItemContainerGenerator.ContainerFromItem($_)
                    if ($container) {
                        $checkbox = $container.ContentTemplate.FindName("ouCheckBox", $container)
                        if (-not $checkbox) {
                            # If named checkbox not found, try to find first checkbox in container
                            $checkbox = [Windows.Media.VisualTreeHelper]::GetChild($container, 0) -as [System.Windows.Controls.CheckBox]
                        }
                        if (-not $checkbox) {
                            Write-Log "Warning: Could not find checkbox for item: $_"
                            return $false
                        }
                        $checkbox.IsChecked
                    }
                    else { 
                        Write-Log "Warning: No container found for item: $_"
                        $false 
                    }
                } | ForEach-Object { $_['DistinguishedName'] }
            
            Write-Log "Found $($selectedOUs.Count) selected OUs"
            
            if (-not $selectedOUs) {
                [System.Windows.MessageBox]::Show("Please select at least one Organizational Unit to analyze.", "No OUs Selected")
                return
            }
            
            Write-Log "Starting analysis of selected OUs..."
            $script:LogOverlay.Visibility = "Visible"
            $script:StopProcessing = $false
            
            # Get group details
            $groups = Get-GroupDetails -SelectedOUs $selectedOUs
            
            if ($groups) {
                Write-Log "Analysis complete. Generating report..."
                New-HTMLReport -Groups $groups
            }
            else {
                Write-Log "Analysis failed or was cancelled."
                [System.Windows.MessageBox]::Show("Analysis failed or was cancelled. Please check the logs for details.", "Error")
            }
        }
        catch {
            Write-Log "Error during report generation: $_"
            [System.Windows.MessageBox]::Show("An error occurred while generating the report. Please check the logs for details.`n`nError: $_", "Error")
        }
    })

    $Window.Add_Closed({
        # Cleanup on window close
        $script:StopProcessing = $true
        $script:LogTextBox = $null
        $script:LogOverlay = $null
        $script:Window = $null
        
        # Clean up any remaining jobs
        Get-Job | Where-Object { $_.Name -eq 'OULoadJob' } | Remove-Job -Force -ErrorAction SilentlyContinue
    })

    # Show the window before loading OUs
    Write-Log "Starting AD Group Review Tool" -Type Success

    # Create a background job to load OUs
    Write-Log "Loading Organizational Units..." -Type Info
    
    # Clean up any existing jobs first
    Get-Job | Where-Object { $_.Name -eq 'OULoadJob' } | Remove-Job -Force -ErrorAction SilentlyContinue
    
    # Create job with version-specific parameters
    $jobParams = @{
        Name = 'OULoadJob'
        ScriptBlock = {
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                $domain = Get-ADDomain -ErrorAction Stop
                $ous = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName -SearchBase $domain.DistinguishedName |
                    ForEach-Object {
                        # Get group count for this OU - use ErrorAction for PS7 compatibility
                        $groupCount = @(Get-ADGroup -Filter * -SearchBase $_.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue).Count
                        
                        # Format the OU path for display
                        $ouPath = $_.DistinguishedName
                        $ouPath = $ouPath -replace '(,DC=[\w-]+)+$', ''
                        $ouPath = $ouPath -replace ',OU=', ' -> '
                        $ouPath = $ouPath -replace '^OU=', ''
                        
                        @{
                            Name = "$ouPath ($groupCount groups)"
                            DistinguishedName = $_.DistinguishedName
                            Description = "Full Path: $($_.DistinguishedName)"
                            GroupCount = $groupCount
                        }
                    } | Sort-Object { $_['GroupCount'] } -Descending
                
                return $ous
            }
            catch {
                throw "Error loading OUs: $_"
            }
        }
    }

    # Add PS7-specific job parameters if needed
    if ($PSVersion -ge 6) {
        $jobParams['WorkingDirectory'] = $PWD.Path
    }

    $OULoadJob = Start-Job @jobParams

    # Create a timer to check job status
    $Timer = New-Object System.Windows.Threading.DispatcherTimer
    $Timer.Interval = [TimeSpan]::FromMilliseconds(500)
    $Timer.Add_Tick({
        if ($OULoadJob.State -eq 'Completed') {
            $Timer.Stop()
            try {
                # Use ErrorAction for PS7 compatibility
                $OUs = Receive-Job -Job $OULoadJob -ErrorAction Stop -Wait
                
                # Clean up the job
                Get-Job | Where-Object { $_.Name -eq 'OULoadJob' } | Remove-Job -Force -ErrorAction SilentlyContinue
                
                if (-not $OUs -or $OUs.Count -eq 0) {
                    Write-Log "No OUs found"
                    $Window.Dispatcher.Invoke({
                        [System.Windows.MessageBox]::Show("No Organizational Units found. Please check your Active Directory connection and try again.", "No OUs Found")
                        $Window.Close()
                    })
                    return
                }
                
                $Window.Dispatcher.Invoke({
                    $OUList.ItemsSource = $OUs
                    $GenerateButton.IsEnabled = $true
                    $ToggleSelectButton.IsEnabled = $true
                    $script:LoadingOverlay.Visibility = "Collapsed"
                    Write-Log "OU list loaded successfully ($($OUs.Count) OUs found)" -Type Success
                })
            }
            catch {
                Write-Log "Error processing OU load results: $_"
                $Window.Dispatcher.Invoke({
                    [System.Windows.MessageBox]::Show("Error loading OUs: $_", "Error")
                    $Window.Close()
                })
            }
        }
        elseif ($OULoadJob.State -eq 'Failed') {
            $Timer.Stop()
            # Use Wait parameter for PS7 compatibility
            $errorMessage = Receive-Job -Job $OULoadJob -ErrorAction SilentlyContinue -Wait
            
            # Clean up the job
            Get-Job | Where-Object { $_.Name -eq 'OULoadJob' } | Remove-Job -Force -ErrorAction SilentlyContinue
            
            Write-Log "Error loading OUs: $errorMessage"
            $Window.Dispatcher.Invoke({
                [System.Windows.MessageBox]::Show("Error loading OUs: $errorMessage", "Error")
                $Window.Close()
            })
        }
    })
    $Timer.Start()

    # Keep the window open and modal
    $Window.ShowDialog()
}
catch {
    Write-Log "Error initializing window: $_"
    [System.Windows.MessageBox]::Show("Error initializing application: $_`n`nPlease check the logs for details.", "Error")
    
    # Clean up any remaining jobs
    Get-Job | Where-Object { $_.Name -eq 'OULoadJob' } | Remove-Job -Force -ErrorAction SilentlyContinue
    return
}