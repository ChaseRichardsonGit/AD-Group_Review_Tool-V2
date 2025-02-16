# Import required modules and assemblies
Import-Module ActiveDirectory
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Collections

# Create log directory if it doesn't exist
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
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

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [switch]$NoConsole
    )
    try {
        $LogMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
        Add-Content -Path $LogFile -Value $LogMessage -ErrorAction Stop
        
        if (-not $NoConsole) {
            Write-Host $LogMessage
        }
        
        if ($script:LogTextBox -and $script:Window) {
            $script:Window.Dispatcher.Invoke(
                [Action]{
                    # Insert new text at the beginning
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
        Write-Host "Error in Write-Log: $_"
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
        Issues = @()
        Score = 100  # Start with perfect score
    }
    
    # Check for empty description
    if ([string]::IsNullOrWhiteSpace($Group.Description)) {
        $health.Issues += "Missing description"
        $health.Score -= 20
    }
    
    # Check for missing manager
    if ([string]::IsNullOrWhiteSpace($Group.Manager)) {
        $health.Issues += "No manager assigned"
        $health.Score -= 20
    }
    
    # Check member count
    if ($Group.TotalMembers -eq 0) {
        $health.Issues += "Empty group"
        $health.Score -= 30
    }
    elseif ($Group.TotalMembers -gt 1000) {
        $health.Issues += "Large group (>1000 members)"
        $health.Score -= 10
    }
    
    # Check age
    $age = (Get-Date) - $Group.Created
    if ($age.Days -gt 365 * 2) {
        $health.Issues += "Group older than 2 years"
        $health.Score -= 10
    }
    
    # Check disabled user percentage
    if ($Group.UserMembers -gt 0) {
        $disabledPercentage = ($Group.DisabledMembers / $Group.UserMembers) * 100
        if ($disabledPercentage -gt 40) {
            $health.Issues += "High percentage of disabled users (>40%)"
            $health.Score -= 30
        }
        elseif ($disabledPercentage -gt 20) {
            $health.Issues += "Moderate percentage of disabled users (>20%)"
            $health.Score -= 15
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
        [System.Collections.Generic.HashSet[string]]$ProcessedGroups = $null
    )
    
    if ($null -eq $ProcessedGroups) {
        $ProcessedGroups = New-Object System.Collections.Generic.HashSet[string]
    }
    
    # If we've already processed this group, skip it to avoid cycles
    if (-not $ProcessedGroups.Add($GroupDN)) {
        return @()
    }
    
    try {
        $group = Get-ADGroup -Identity $GroupDN -Properties memberOf
        $nestedGroups = @()
        
        foreach ($memberOfGroup in $group.memberOf) {
            if ($ProcessedGroups.Contains($memberOfGroup)) {
                continue  # Skip if already processed to avoid cycles
            }
            
            $nestedGroups += $memberOfGroup
            # Recursively get nested groups
            $nestedGroups += Get-NestedGroupMembership -GroupDN $memberOfGroup -ProcessedGroups $ProcessedGroups
        }
        
        return $nestedGroups | Select-Object -Unique
    }
    catch {
        Write-Log "Error getting nested groups for ${GroupDN}: $_" -NoConsole
        return @()
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
        
        $allGroups = @()
        $processedGroups = 0
        $totalGroups = 0
        
        # Create hashtable to store OU statistics
        ${script:OUStats} = @{}
        
        # First pass - count total groups
        foreach($ou in $SelectedOUs) {
            $groups = Get-ADGroup -Filter * -SearchBase $ou
            $totalGroups += $groups.Count
            
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
                    $nestingDepth = ($nestedGroups | Measure-Object).Count
                    
                    # Update OU statistics for nested groups
                    ${script:OUStats}[$ou].NestedGroupCount += $nestingDepth
                    ${script:OUStats}[$ou].MaxNestingDepth = [Math]::Max(${script:OUStats}[$ou].MaxNestingDepth, $nestingDepth)
                    
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
                        $userMembers = $users.Count
                        $enabledMembers = ($users | Where-Object { $_.Enabled }).Count
                        $disabledMembers = ($users | Where-Object { -not $_.Enabled }).Count
                        
                        Start-Sleep -Milliseconds 50
                        
                        $groupMembers = @(Get-ADGroup -LDAPFilter "(memberOf=$($group.DistinguishedName))" -ResultSetSize $null).Count
                        Start-Sleep -Milliseconds 50
                        
                        $computerMembers = @(Get-ADComputer -LDAPFilter "(memberOf=$($group.DistinguishedName))" -ResultSetSize $null).Count
                        Start-Sleep -Milliseconds 50
                        
                        $totalMembers = $userMembers + $groupMembers + $computerMembers
                        
                        # Update OU statistics
                        ${script:OUStats}[$ou].EnabledMembers += $enabledMembers
                        ${script:OUStats}[$ou].DisabledMembers += $disabledMembers
                        ${script:OUStats}[$ou].TotalMembers += $totalMembers
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
                        NestedInGroupCount = @($group.memberOf).Count
                        HasNestedGroups = ($groupMembers -gt 0)
                        DN = $group.DistinguishedName
                        OU = ($group.DistinguishedName -split ',',2)[1]
                        SamAccountName = $group.sAMAccountName
                    }
                    
                    # Add health check
                    $health = Get-GroupHealth $groupObj
                    $groupObj | Add-Member -NotePropertyName HealthScore -NotePropertyValue $health.Score
                    $groupObj | Add-Member -NotePropertyName HealthIssues -NotePropertyValue $health.Issues
                    
                    $allGroups += $groupObj
                }
                catch {
                    Write-Log "Error processing group $($group.Name): $_"
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
        $emptyGroups = @($Groups | Where-Object { $_.TotalMembers -eq 0 }).Count
        $noManager = @($Groups | Where-Object { -not $_.Manager }).Count
        $noDescription = @($Groups | Where-Object { -not $_.Description }).Count
        $nestedGroups = @($Groups | Where-Object { $_.HasNestedGroups }).Count
        $avgHealth = ($Groups | Measure-Object -Property HealthScore -Average).Average
        $criticalGroups = @($Groups | Where-Object { $_.HealthScore -le 50 }).Count
        $warningGroups = @($Groups | Where-Object { $_.HealthScore -gt 50 -and $_.HealthScore -le 80 }).Count
        $healthyGroups = @($Groups | Where-Object { $_.HealthScore -gt 80 }).Count
        
        # Process OU statistics
        $ouStats = ${script:OUStats}.GetEnumerator() | ForEach-Object {
            # Clean up OU name by removing OU= prefixes, DC components, and improving path readability
            $ouName = $_.Key -replace '(,DC=[\w-]+)+$', ''  # Remove DC components
            $ouName = $ouName -replace '^OU=|,OU=', ''      # Remove all OU= prefixes
            
            # Split path components and reverse them to get most specific first
            $parts = $ouName -split ','
            # Take only unique parts to avoid redundancy
            $uniqueParts = $parts | Select-Object -Unique
            # Join back together with arrows, removing any empty parts
            $ouName = ($uniqueParts | Where-Object { $_ -match '\S' }) -join ' -> '
            
            $stats = $_.Value
            $disabledPercentage = if ($stats.TotalMembers -gt 0) {
                [math]::Round(($stats.DisabledMembers / $stats.TotalMembers) * 100, 1)
            } else { 0 }
            
            @{
                OU = $ouName
                Count = $stats.GroupCount
                EnabledMembers = $stats.EnabledMembers
                DisabledMembers = $stats.DisabledMembers
                DisabledPercentage = $disabledPercentage
            }
        } | Sort-Object { $_.Count } -Descending

        # Generate HTML with enhanced styling
        $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AD Group Health Report</title>
    <style>
        :root {
            --primary-color: #007ACC;
            --warning-color: #dd6b20;
            --critical-color: #e53e3e;
            --success-color: #2f855a;
            --bg-color: #f8f9fa;
        }
        
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: var(--bg-color);
            color: #333;
            line-height: 1.6;
        }
        
        .container {
            max-width: 1600px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 16px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        }

        h1 {
            font-size: 32px;
            font-weight: 600;
            color: #1a202c;
            margin-bottom: 32px;
            padding-bottom: 16px;
            border-bottom: 2px solid #edf2f7;
        }

        .dashboard {
            display: grid;
            grid-template-columns: 1.2fr 0.8fr 1fr 1fr;
            gap: 24px;
            margin-bottom: 40px;
        }

        .card {
            background: white;
            padding: 24px;
            border-radius: 12px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.04);
            border: 1px solid #edf2f7;
        }

        .card h2 {
            font-size: 18px;
            font-weight: 600;
            color: #2d3748;
            margin: 0 0 20px 0;
            padding-bottom: 12px;
            border-bottom: 1px solid #edf2f7;
        }

        /* Overview Card Styles */
        .stat {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
        }

        .stat-label {
            font-size: 14px;
            color: #4a5568;
        }

        .stat-value {
            font-size: 24px;
            font-weight: 600;
            color: #2d3748;
        }

        .stat-value.health {
            color: var(--primary-color);
        }

        .progress-bar {
            height: 8px;
            background: #edf2f7;
            border-radius: 4px;
            overflow: hidden;
            margin-top: 16px;
        }

        .progress-fill {
            height: 100%;
            background: var(--primary-color);
            border-radius: 4px;
            transition: width 0.3s ease;
        }

        /* Health Distribution Styles */
        .health-distribution {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .health-segment {
            padding: 16px;
            border-radius: 8px;
            text-align: center;
            color: white;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .health-segment div {
            font-size: 16px;
            font-weight: 500;
            margin: 0;
        }

        .health-segment strong {
            font-size: 28px;
            font-weight: 600;
            margin-left: 16px;
        }

        .critical {
            background: var(--critical-color);
        }

        .warning {
            background: var(--warning-color);
        }

        .healthy {
            background: var(--success-color);
        }

        /* OU Statistics Styles */
        .ou-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 16px;
        }

        .ou-card {
            background: #f8fafc;
            border-radius: 8px;
            padding: 20px;
            text-align: center;
            transition: transform 0.2s, box-shadow 0.2s;
            border: 1px solid #edf2f7;
        }

        .ou-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }

        .ou-name {
            font-size: 16px;
            font-weight: 600;
            color: #2d3748;
            margin-bottom: 8px;
        }

        .ou-count {
            font-size: 14px;
            color: #718096;
            margin-bottom: 12px;
        }

        .ou-health {
            font-size: 24px;
            font-weight: 600;
            color: var(--primary-color);
        }

        /* Issues Overview Styles */
        .stat-icon {
            width: 32px;
            height: 32px;
            padding: 6px;
            border-radius: 8px;
            margin-right: 12px;
        }

        .stat-icon.empty {
            background: #fed7d7;
            color: var(--critical-color);
        }

        .stat-icon.manager {
            background: #feebc8;
            color: var(--warning-color);
        }

        .stat-icon.description {
            background: #e9d8fd;
            color: #6b46c1;
        }

        .stat-icon.nested {
            background: #c6f6d5;
            color: var(--success-color);
        }

        /* Filter Section Styles */
        .filter-section {
            display: flex;
            gap: 12px;
            margin: 24px 0;
        }

        .filter-button {
            padding: 10px 20px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            background: #edf2f7;
            color: #4a5568;
            transition: all 0.2s;
        }

        .filter-button:hover {
            background: #e2e8f0;
        }

        .filter-button.active {
            background: var(--primary-color);
            color: white;
        }

        /* Table Styles */
        table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin: 24px 0;
            background: white;
            border-radius: 8px;
            overflow: hidden;
        }

        th {
            background: #f8fafc;
            padding: 16px;
            text-align: left;
            font-weight: 600;
            color: #2D3748;
            border-bottom: 2px solid #e2e8f0;
            white-space: nowrap;
        }

        td {
            padding: 16px;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: top;
        }

        /* Column Widths */
        th:nth-child(1), td:nth-child(1) { width: 20%; }  /* Group Name */
        th:nth-child(2), td:nth-child(2) { width: 8%; }   /* Health */
        th:nth-child(3), td:nth-child(3) { width: 15%; }  /* Members */
        th:nth-child(4), td:nth-child(4) { width: 25%; }  /* Description */
        th:nth-child(5), td:nth-child(5) { width: 17%; }  /* Manager */
        th:nth-child(6), td:nth-child(6) { width: 15%; }  /* Details */

        tr:hover {
            background-color: #f7fafc;
        }

        /* Health Badge Styles */
        .badge {
            display: inline-flex;
            align-items: center;
            padding: 6px 12px;
            border-radius: 16px;
            font-weight: 600;
            font-size: 0.875rem;
            line-height: 1;
            white-space: nowrap;
        }

        .badge::before {
            content: '';
            display: inline-block;
            width: 8px;
            height: 8px;
            border-radius: 50%;
            margin-right: 6px;
        }

        .badge-critical {
            color: var(--critical-color);
            background: #fff5f5;
        }
        .badge-critical::before {
            background: var(--critical-color);
        }

        .badge-warning {
            color: var(--warning-color);
            background: #fffaf0;
        }
        .badge-warning::before {
            background: var(--warning-color);
        }

        .badge-success {
            color: var(--success-color);
            background: #f0fff4;
        }
        .badge-success::before {
            background: var(--success-color);
        }

        /* Member Stats Styles */
        .member-stats {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .member-stat {
            display: inline-flex;
            align-items: center;
            font-size: 0.875rem;
            color: #4a5568;
        }

        /* Issues List Styles */
        .issues-list {
            margin-top: 8px;
            font-size: 0.875rem;
            color: #718096;
        }

        .issues-list > * {
            display: flex;
            align-items: center;
            gap: 4px;
        }

        .issues-list > *::before {
            content: '•';
            color: #cbd5e0;
        }

        /* Manager Info Styles */
        .manager-info {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .manager-name {
            font-weight: 600;
            color: #2d3748;
        }

        .manager-title {
            color: #718096;
            font-size: 0.875rem;
        }

        .health-critical {
            color: var(--critical-color);
            font-weight: 500;
        }

        /* Details Section Styles */
        td > div {
            margin-bottom: 4px;
        }

        .nested-warning {
            color: var(--warning-color);
        }

        /* OU Statistics Table Styles */
        .ou-stats-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 16px;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }

        .ou-stats-table th,
        .ou-stats-table td {
            padding: 8px 12px;
            text-align: left;
            border-bottom: 1px solid #e2e8f0;
            font-size: 13px;
        }

        .ou-stats-table th {
            background: #f8fafc;
            font-weight: 600;
            color: #2d3748;
            text-align: center;
        }

        .ou-stats-table td {
            vertical-align: middle;
        }

        .ou-stats-table th.no-bottom-border {
            border-bottom: none;
            padding-bottom: 4px;
        }

        .ou-stats-table th.top-border {
            border-top: 1px solid #e2e8f0;
            padding-top: 4px;
        }

        .members-cell {
            text-align: center;
            white-space: nowrap;
        }

        .text-center {
            text-align: center;
        }

        .warning-text {
            color: var(--warning-color);
            font-weight: 500;
        }

        .critical-text {
            color: var(--critical-color);
            font-weight: 500;
        }

        .ou-stats-table td:first-child {
            max-width: 300px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
    </style>
    <script>
        function filterGroups(minHealth, maxHealth) {
            const rows = document.querySelectorAll('table tr:not(:first-child)');
            rows.forEach(row => {
                const healthScore = parseInt(row.querySelector('.health-score').textContent);
                row.style.display = (healthScore >= minHealth && healthScore <= maxHealth) ? '' : 'none';
            });
            
            document.querySelectorAll('.filter-button').forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');
        }
        
        function searchGroups() {
            const searchText = document.getElementById('groupSearch').value.toLowerCase();
            const rows = document.querySelectorAll('table tr:not(:first-child)');
            
            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                row.style.display = text.includes(searchText) ? '' : 'none';
            });
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>AD Group Health Report</h1>
        <div class="dashboard">
            <!-- Overview Card -->
            <div class="card">
                <h2>Overview</h2>
                <div class="stat">
                    <div class="stat-info">
                        <span class="stat-label">Total Groups</span>
                        <span class="stat-value">$totalGroups</span>
                    </div>
                    <svg class="stat-icon" viewBox="0 0 24 24">
                        <path fill="currentColor" d="M12 7V3H2v18h20V7H12zM6 19H4v-2h2v2zm0-4H4v-2h2v2zm0-4H4V9h2v2zm0-4H4V5h2v2zm4 12H8v-2h2v2zm0-4H8v-2h2v2zm0-4H8V9h2v2zm0-4H8V5h2v2z"/>
                    </svg>
                </div>
                <div class="stat">
                    <div class="stat-info">
                        <span class="stat-label">Average Health</span>
                        <span class="stat-value health">$([Math]::Round($avgHealth, 1))%</span>
                    </div>
                    <svg class="stat-icon" viewBox="0 0 24 24">
                        <path fill="currentColor" d="M19.03 7.39l1.42-1.42c-.43-.51-.9-.99-1.41-1.41l-1.42 1.42C16.07 4.74 14.12 4 12 4c-4.97 0-9 4.03-9 9s4.02 9 9 9 9-4.48 9-9c0-2.12-.74-4.07-1.97-5.61zM13 14h-2V8h2v6z"/>
                    </svg>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: $([Math]::Round($avgHealth, 1))%;"></div>
                </div>
            </div>
            
            <!-- Health Distribution Card -->
            <div class="card">
                <h2>Health Distribution</h2>
                <div class="health-distribution">
                    <div class="health-segment critical">
                        <div>Critical</div>
                        <strong>$criticalGroups</strong>
                    </div>
                    <div class="health-segment warning">
                        <div>Warning</div>
                        <strong>$warningGroups</strong>
                    </div>
                    <div class="health-segment healthy">
                        <div>Healthy</div>
                        <strong>$healthyGroups</strong>
                    </div>
                </div>
            </div>
            
            <!-- Issues Overview Card -->
            <div class="card">
                <h2>Issues Overview</h2>
                <div class="stat">
                    <div class="stat-info">
                        <span class="stat-label">Empty Groups</span>
                        <span class="stat-value">$emptyGroups</span>
                    </div>
                    <svg class="stat-icon empty" viewBox="0 0 24 24">
                        <path fill="currentColor" d="M19 5v14H5V5h14m0-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zm-7-2h2V7h-4v2h2z"/>
                    </svg>
                </div>
                <div class="stat">
                    <div class="stat-info">
                        <span class="stat-label">No Manager</span>
                        <span class="stat-value">$noManager</span>
                    </div>
                    <svg class="stat-icon manager" viewBox="0 0 24 24">
                        <path fill="currentColor" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 3c1.66 0 3 1.34 3 3s-1.34 3-3 3-3-1.34-3-3 1.34-3 3-3zm0 14.2c-2.5 0-4.71-1.28-6-3.22.03-1.99 4-3.08 6-3.08 1.99 0 5.97 1.09 6 3.08-1.29 1.94-3.5 3.22-6 3.22z"/>
                    </svg>
                </div>
                <div class="stat">
                    <div class="stat-info">
                        <span class="stat-label">No Description</span>
                        <span class="stat-value">$noDescription</span>
                    </div>
                    <svg class="stat-icon description" viewBox="0 0 24 24">
                        <path fill="currentColor" d="M14 17H4v2h10v-2zm6-8H4v2h16V9zM4 15h16v-2H4v2zM4 5v2h16V5H4z"/>
                    </svg>
                </div>
                <div class="stat">
                    <div class="stat-info">
                        <span class="stat-label">Nested Groups</span>
                        <span class="stat-value">$nestedGroups</span>
                    </div>
                    <svg class="stat-icon nested" viewBox="0 0 24 24">
                        <path fill="currentColor" d="M3 5v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2zm16 14H5V5h14v14zm-7-2h2V7h-4v2h2z"/>
                    </svg>
                </div>
            </div>
            
            <!-- OU Statistics Card -->
            <div class="card">
                <h2>OU Statistics</h2>
                <table class="ou-stats-table">
                    <thead>
                        <tr>
                            <th rowspan="2" style="text-align: left;">Organizational Unit</th>
                            <th rowspan="2" style="text-align: center;">Groups</th>
                            <th colspan="2" class="no-bottom-border">Members</th>
                            <th colspan="2" class="no-bottom-border">Nesting</th>
                        </tr>
                        <tr>
                            <th class="top-border">Enabled / Disabled</th>
                            <th class="top-border">% Disabled</th>
                            <th class="top-border">Total Nested</th>
                            <th class="top-border">Max Depth</th>
                        </tr>
                    </thead>
                    <tbody>
                        $(foreach ($stat in $ouStats) {
                            $disabledClass = if ($stat.DisabledPercentage -gt 20) { 'warning-text' } elseif ($stat.DisabledPercentage -gt 40) { 'critical-text' } else { '' }
                            $nestingClass = if ($stat.MaxNestingDepth -gt 5) { 'warning-text' } elseif ($stat.MaxNestingDepth -gt 10) { 'critical-text' } else { '' }
                            @"
                            <tr>
                                <td>$($stat.OU)</td>
                                <td style="text-align: center;">$($stat.Count)</td>
                                <td class="members-cell">$($stat.EnabledMembers) / $($stat.DisabledMembers)</td>
                                <td class="text-center $disabledClass">$($stat.DisabledPercentage)%</td>
                                <td class="text-center">$($stat.NestedGroupCount)</td>
                                <td class="text-center $nestingClass">$($stat.MaxNestingDepth)</td>
                            </tr>
"@
                        })
                    </tbody>
                </table>
            </div>
        </div>
        
        <input type="text" id="groupSearch" class="search-box" placeholder="Search groups..." onkeyup="searchGroups()">
        
        <div class="filter-section">
            <button class="filter-button active" onclick="filterGroups(0, 100)">All Groups</button>
            <button class="filter-button" onclick="filterGroups(0, 50)">Critical (0-50)</button>
            <button class="filter-button" onclick="filterGroups(51, 80)">Warning (51-80)</button>
            <button class="filter-button" onclick="filterGroups(81, 100)">Healthy (81-100)</button>
        </div>

        <table>
            <tr>
                <th>Group Name</th>
                <th>Health</th>
                <th>Members</th>
                <th>Description & Issues</th>
                <th>Manager</th>
                <th>Details</th>
            </tr>
"@

        # Add rows for each group
        # Remove duplicates and sort by member count (descending) and then name
        $sortedGroups = $Groups | Sort-Object -Unique Name | Sort-Object @{Expression={$_.TotalMembers}; Descending=$true}, Name
        foreach ($group in $sortedGroups) {
            $healthBadge = if ($group.HealthScore -le 50) {
                'badge-critical'
            } elseif ($group.HealthScore -le 80) {
                'badge-warning'
            } else {
                'badge-success'
            }
            
            $HTML += @"
            <tr>
                <td>
                    <strong>$($group.Name)</strong>
                    <div class="group-details">
                        <span>$($group.SamAccountName)</span>
                        <span>$($group.OU)</span>
                    </div>
                </td>
                <td>
                    <span class="badge $healthBadge health-score">$($group.HealthScore)</span>
                </td>
                <td>
                    <div class="member-stats">
                        <span class="member-stat">
                            <svg viewBox="0 0 24 24" width="16" height="16" style="margin-right: 4px;">
                                <path fill="currentColor" d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/>
                            </svg>
                            Total: $($group.TotalMembers)
                        </span>
                        <span class="member-stat">
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5z"/>
                            </svg>
                            Users: $($group.UserMembers)
                        </span>
                        <span class="member-stat">
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M12 7V3H2v18h20V7H12zM6 19H4v-2h2v2zm0-4H4v-2h2v2zm0-4H4V9h2v2zm0-4H4V5h2v2zm4 12H8v-2h2v2zm0-4H8v-2h2v2zm0-4H8V9h2v2zm0-4H8V5h2v2zm10 12h-8v-2h2v-2h-2v-2h2v-2h-2V9h8v10zm-2-8h-2v2h2v-2z"/>
                            </svg>
                            Groups: $($group.GroupMembers)
                        </span>
                        <span class="member-stat">
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M21 2H3c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h7l-2 3v1h8v-1l-2-3h7c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z"/>
                            </svg>
                            Computers: $($group.ComputerMembers)
                        </span>
                    </div>
                </td>
                <td>
                    <div class="description">$($group.Description)</div>
                    $(if ($group.Info) {"<div class='notes'>$($group.Info)</div>"})
                    $(if ($group.HealthIssues) {
                        "<div class='issues-list'>" + 
                        ($group.HealthIssues | ForEach-Object { 
                            "<div class='issue-item'>$([System.Web.HttpUtility]::HtmlEncode($_))</div>" 
                        }) -join "`n" +
                        "</div>"
                    })
                </td>
                <td>
                    $(if ($group.Manager) {
                        @"
                        <div class="manager-info">
                            <span class="manager-name">$([System.Web.HttpUtility]::HtmlEncode($group.Manager.DisplayName))</span>
                            <span class="manager-title">$([System.Web.HttpUtility]::HtmlEncode($group.Manager.Title))</span>
                            <span class="manager-upn">$([System.Web.HttpUtility]::HtmlEncode($group.Manager.UPN))</span>
                        </div>
"@
                    } else {
                        "<div class='health-critical'>
                            <svg viewBox='0 0 24 24' width='16' height='16' style='margin-right: 4px;'>
                                <path fill='currentColor' d='M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z'/>
                            </svg>
                            No manager assigned
                        </div>"
                    })
                </td>
                <td>
                    <div class="details-section">
                        <div>
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm0 10.99h7c-.53 4.12-3.28 7.79-7 8.94V12H5V6.3l7-3.11v8.8z"/>
                            </svg>
                            Category: $($group.Category)
                        </div>
                        <div>
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8z"/>
                            </svg>
                            Scope: $($group.Scope)
                        </div>
                        <div>
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.11 0-1.99.9-1.99 2L3 19c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z"/>
                            </svg>
                            Created: $($group.Created.ToString('yyyy-MM-dd'))
                        </div>
                        $(if ($group.HasNestedGroups -or $group.NestedInGroupCount -gt 0) {
                            "<div class='nested-warning'>
                                Nested in $($group.NestedInGroupCount) groups" +
                                $(if ($group.HasNestedGroups) { " / Has nested members" }) +
                            "</div>"
                        })
                    </div>
                </td>
            </tr>
"@
        }

        $HTML += "</table></div></body></html>"
        
        # Save report with UTF8 encoding without BOM
        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
        [System.IO.File]::WriteAllLines($ReportFile, $HTML, $Utf8NoBomEncoding)
        Write-Log "Report saved to: $ReportFile"
        
        # Open the downloads folder and report in a UI-safe way
        $script:Window.Dispatcher.Invoke({
            try {
                Write-Log "Opening report location..."
                # Open folder and select both files
                $files = @($ReportFile, $CSVFile)
                $filesArg = $files -join '" "'  # Join paths with quotes
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

# Define the XAML for the WPF GUI
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AD Group Review Tool" 
    Height="900" 
    Width="1200"
    WindowStartupLocation="CenterScreen"
    Background="#f0f2f5">
    <Grid>
        <!-- Loading Overlay -->
        <Border x:Name="loadingOverlay" 
                Background="#80000000" 
                Visibility="Visible"
                Panel.ZIndex="1000">
            <Border Background="White" 
                    CornerRadius="12" 
                    Width="400"
                    Height="200"
                    VerticalAlignment="Center">
                <Grid Margin="24">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Text="Please wait, gathering list of OUs and groups in each OU..." 
                             TextWrapping="Wrap"
                             FontSize="16"
                             FontWeight="SemiBold"
                             HorizontalAlignment="Center"
                             Margin="0,0,0,20"/>
                    
                    <ProgressBar Grid.Row="1" 
                               IsIndeterminate="True" 
                               Height="4" 
                               Background="Transparent"
                               Foreground="#007ACC"/>
                </Grid>
            </Border>
        </Border>

        <!-- Main Content -->
        <Border Margin="24" Background="White" CornerRadius="20" Padding="32">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Header with Icon -->
                <StackPanel Grid.Row="0" Margin="0,0,0,20" Orientation="Horizontal">
                    <Viewbox Width="32" Height="32" Margin="0,0,16,0">
                        <Path Data="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm0-14c-3.31 0-6 2.69-6 6s2.69 6 6 6 6-2.69 6-6-2.69-6-6-6zm0 10c-2.21 0-4-1.79-4-4s1.79-4 4-4 4 1.79 4 4-1.79 4-4 4z"
                              Fill="#007ACC"/>
                    </Viewbox>
                    <StackPanel>
                        <TextBlock Text="AD Group Review Tool" 
                                 FontSize="28" 
                                 FontWeight="SemiBold" 
                                 Foreground="#2D3748"/>
                        <TextBlock Text="Analyze and optimize Active Directory groups" 
                                 FontSize="14" 
                                 Foreground="#718096"/>
                    </StackPanel>
                </StackPanel>

                <!-- Main Content Area -->
                <Grid Grid.Row="1">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Feature Icons -->
                    <UniformGrid Grid.Row="0" Rows="1" Margin="0,0,0,20">
                        <Border Background="#f8fafc" CornerRadius="8" Padding="16" Margin="4">
                            <StackPanel>
                                <Viewbox Width="24" Height="24">
                                    <Path Data="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.11 0 2-.9 2-2V5c0-1.1-.89-2-2-2zm-9 14l-5-5 1.41-1.41L10 14.17L17.59 6.58L19 8l-9 9zM1 9h4v12H1z" 
                                          Fill="#4A5568"/>
                                </Viewbox>
                                <TextBlock Text="Group Analysis" 
                                         FontSize="14" 
                                         FontWeight="SemiBold"
                                         Foreground="#2D3748" 
                                         HorizontalAlignment="Center" 
                                         Margin="0,8,0,4"/>
                                <TextBlock Text="Comprehensive analysis of AD group structure and membership"
                                         TextAlignment="Center"
                                         TextWrapping="Wrap"
                                         FontSize="12"
                                         Foreground="#718096"
                                         Margin="0,0,0,8"/>
                            </StackPanel>
                        </Border>
                        <Border Background="#f8fafc" CornerRadius="8" Padding="16" Margin="4">
                            <StackPanel>
                                <Viewbox Width="24" Height="24">
                                    <Path Data="M9 21h9c.83 0 1.54-.5 1.84-1.22l3.02-7.05c.09-.23.14-.47.14-.73v-2c0-1.1-.9-2-2-2h-6.31l.95-4.57.03-.32c0-.41-.17-.79-.44-1.06L14.17 1 7.58 7.59C7.22 7.95 7 8.45 7 9v10c0 1.1.9 2 2 2zM9 9l4.34-4.34L12 10h9v2l-3 7H9V9zM1 9h4v12H1z" 
                                          Fill="#4A5568"/>
                                </Viewbox>
                                <TextBlock Text="Health Check" 
                                         FontSize="14"
                                         FontWeight="SemiBold" 
                                         Foreground="#2D3748" 
                                         HorizontalAlignment="Center" 
                                         Margin="0,8,0,4"/>
                                <TextBlock Text="Evaluate group health scores and identify potential issues"
                                         TextAlignment="Center"
                                         TextWrapping="Wrap"
                                         FontSize="12"
                                         Foreground="#718096"
                                         Margin="0,0,0,8"/>
                            </StackPanel>
                        </Border>
                        <Border Background="#f8fafc" CornerRadius="8" Padding="16" Margin="4">
                            <StackPanel>
                                <Viewbox Width="24" Height="24">
                                    <Path Data="M12 8c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4zm8.94 3c-.46-4.17-3.77-7.48-7.94-7.94V1h-2v2.06C6.83 3.52 3.52 6.83 3.06 11H1v2h2.06c.46 4.17 3.77 7.48 7.94 7.94V23h2v-2.06c4.17-.46 7.48-3.77 7.94-7.94H23v-2h-2.06zM12 19c-3.87 0-7-3.13-7-7s3.13-7 7-7 7 3.13 7 7-3.13 7-7 7z" 
                                          Fill="#4A5568"/>
                                </Viewbox>
                                <TextBlock Text="Optimization" 
                                         FontSize="14"
                                         FontWeight="SemiBold" 
                                         Foreground="#2D3748" 
                                         HorizontalAlignment="Center" 
                                         Margin="0,8,0,4"/>
                                <TextBlock Text="Recommendations for improving group structure and management"
                                         TextAlignment="Center"
                                         TextWrapping="Wrap"
                                         FontSize="12"
                                         Foreground="#718096"
                                         Margin="0,0,0,8"/>
                            </StackPanel>
                        </Border>
                    </UniformGrid>

                    <!-- OU Selection -->
                    <Border Grid.Row="1" 
                            Background="#f8fafc" 
                            CornerRadius="12" 
                            Padding="20">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,12">
                                <Viewbox Width="20" Height="20" Margin="0,0,8,0">
                                    <Path Data="M3 3h18v18H3z" Fill="#007ACC"/>
                                </Viewbox>
                                <TextBlock Text="Select Organizational Units:" 
                                         FontSize="16" 
                                         FontWeight="SemiBold" 
                                         Foreground="#2D3748"/>
                            </StackPanel>
                            
                            <ScrollViewer Grid.Row="1" 
                                        MaxHeight="240" 
                                        VerticalScrollBarVisibility="Auto"
                                        HorizontalScrollBarVisibility="Disabled">
                                <ItemsControl x:Name="OUList">
                                    <ItemsControl.ItemTemplate>
                                        <DataTemplate>
                                            <CheckBox Margin="0,4,0,4"
                                                    Content="{Binding Name}"
                                                    Tag="{Binding DistinguishedName}"
                                                    ToolTip="{Binding Description}"
                                                    IsChecked="False"
                                                    x:Name="ouCheckBox"
                                                    Foreground="#4A5568"/>
                                        </DataTemplate>
                                    </ItemsControl.ItemTemplate>
                                </ItemsControl>
                            </ScrollViewer>
                        </Grid>
                    </Border>
                </Grid>

                <!-- Generate Button and Toggle Button -->
                <Grid Grid.Row="2" Margin="0,20,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Toggle Select All Button -->
                    <ToggleButton x:Name="btnToggleSelect"
                                Height="44"
                                Padding="20,0"
                                Margin="0,0,10,0">
                        <ToggleButton.Template>
                            <ControlTemplate TargetType="ToggleButton">
                                <Border Background="#f0f2f5" 
                                        CornerRadius="8" 
                                        BorderThickness="1"
                                        BorderBrush="#CBD5E0">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Viewbox Width="20" Height="20" Margin="0,0,8,0">
                                            <Path x:Name="checkIcon"
                                                  Fill="#4A5568"
                                                  Data="M19 3H5C3.89 3 3 3.9 3 5V19C3 20.1 3.89 21 5 21H19C20.11 21 21 20.1 21 19V5C21 3.9 20.11 3 19 3ZM10 17L5 12L6.41 10.59L10 14.17L17.59 6.58L19 8L10 17Z"/>
                                        </Viewbox>
                                        <TextBlock Grid.Column="1" 
                                                 x:Name="toggleText"
                                                 Text="Select All" 
                                                 FontSize="14"
                                                 FontWeight="SemiBold"
                                                 Foreground="#4A5568"
                                                 VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsChecked" Value="True">
                                        <Setter TargetName="toggleText" Property="Text" Value="Deselect All"/>
                                        <Setter TargetName="checkIcon" Property="Data" Value="M19 3H5C3.89 3 3 3.9 3 5V19C3 20.1 3.89 21 5 21H19C20.11 21 21 20.1 21 19V5C21 3.9 20.11 3 19 3ZM16 13H8V11H16V13Z"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </ToggleButton.Template>
                    </ToggleButton>

                    <!-- Generate Button -->
                    <Button Grid.Column="1"
                            x:Name="btnGenerate" 
                            Height="44"
                            Width="200"
                            HorizontalAlignment="Center">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border Background="#007ACC" 
                                        CornerRadius="8" 
                                        Padding="20,0">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <Viewbox Width="20" Height="20" Margin="0,0,8,0">
                                            <Path Data="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z" 
                                                  Fill="White"/>
                                        </Viewbox>
                                        <TextBlock Grid.Column="1" 
                                                 Text="Generate Report" 
                                                 FontSize="16"
                                                 FontWeight="SemiBold"
                                                 Foreground="White"
                                                 VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                </Grid>
            </Grid>
        </Border>

        <!-- Log Overlay -->
        <Border x:Name="logOverlay" 
                Background="#80000000" 
                Visibility="Collapsed">
            <Border Background="White" 
                    CornerRadius="12" 
                    Margin="48" 
                    Padding="24">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                        <Viewbox Width="24" Height="24" Margin="0,0,12,0">
                            <Path Data="M4 4L20 20M4 20L20 4" Fill="Transparent" Stroke="#007ACC"/>
                        </Viewbox>
                        <TextBlock Text="Operation Progress" 
                                 FontSize="20" 
                                 FontWeight="SemiBold"/>
                    </StackPanel>

                    <ScrollViewer Grid.Row="1" 
                                VerticalScrollBarVisibility="Auto">
                        <TextBox x:Name="logTextBox" 
                                IsReadOnly="True" 
                                Background="Transparent" 
                                BorderThickness="0" 
                                FontFamily="Consolas" 
                                FontSize="13"
                                TextWrapping="Wrap"/>
                    </ScrollViewer>

                    <Button Grid.Row="2" 
                            x:Name="btnCloseLog" 
                            Content="Close" 
                            HorizontalAlignment="Right" 
                            Margin="0,16,0,0" 
                            Padding="24,8" 
                            Background="#007ACC" 
                            Foreground="White" 
                            BorderThickness="0">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="6"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                </Grid>
            </Border>
        </Border>
    </Grid>
</Window>
"@

# Create and show the window immediately
try {
    $Reader = [System.Xml.XmlNodeReader]::New($XAML)
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
    Write-Log "Starting AD Group Review Tool"

    # Create a background job to load OUs
    Write-Log "Loading Organizational Units..."
    
    # Clean up any existing jobs first
    Get-Job | Where-Object { $_.Name -eq 'OULoadJob' } | Remove-Job -Force -ErrorAction SilentlyContinue
    
    $OULoadJob = Start-Job -Name 'OULoadJob' -ScriptBlock {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            $domain = Get-ADDomain -ErrorAction Stop
            $ous = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName -SearchBase $domain.DistinguishedName |
                ForEach-Object {
                    # Get group count for this OU
                    $groupCount = @(Get-ADGroup -Filter * -SearchBase $_.DistinguishedName -SearchScope OneLevel).Count
                    
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

    # Create a timer to check job status
    $Timer = New-Object System.Windows.Threading.DispatcherTimer
    $Timer.Interval = [TimeSpan]::FromMilliseconds(500)
    $Timer.Add_Tick({
        if ($OULoadJob.State -eq 'Completed') {
            $Timer.Stop()
            try {
                $OUs = Receive-Job -Job $OULoadJob -ErrorAction Stop
                
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
                    Write-Log "OU list loaded successfully ($($OUs.Count) OUs found)"
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
            $errorMessage = Receive-Job -Job $OULoadJob -ErrorAction SilentlyContinue
            
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