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
function Generate-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$ReportData
    )
    
    try {
        # Get the template path relative to script location
        $templatePath = Join-Path $ScriptDir "templates\report_template.html"
        
        # Read the template
        $template = Get-Content -Path $templatePath -Raw
        
        # Convert report data to JSON for JavaScript
        $jsonData = $ReportData | ConvertTo-Json -Depth 10
        
        # Replace template data placeholder with actual JSON data
        $template = $template -replace 'let templateData = \{\};', "let templateData = $jsonData;"
        
        # Generate output filename with timestamp
        $outputFile = Join-Path $ScriptDir "reports\ADGroupHealth_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        # Ensure reports directory exists
        $reportsDir = Join-Path $ScriptDir "reports"
        if (-not (Test-Path $reportsDir)) {
            New-Item -ItemType Directory -Path $reportsDir | Out-Null
        }
        
        # Save the report
        $template | Out-File -FilePath $outputFile -Encoding UTF8
        
        Write-Log -Message "HTML report generated successfully: $outputFile" -Type Success
        return $outputFile
    }
    catch {
        Write-Log -Message "Error generating HTML report: $_" -Type Error
        return $null
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
                                Width="160"
                                Padding="20,0"
                                Margin="0,0,10,0">
                        <ToggleButton.Template>
                            <ControlTemplate TargetType="ToggleButton">
                                <Border Background="#f0f2f5" 
                                        CornerRadius="8" 
                                        BorderThickness="1"
                                        BorderBrush="#CBD5E0">
                                    <Grid HorizontalAlignment="Center">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="Auto"/>
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