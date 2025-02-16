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

# Set up paths for resource files
$script:XamlFile = Join-Path $PSScriptRoot "Resources\GUI.xaml"
$script:HtmlTemplateFile = Join-Path $PSScriptRoot "Resources\HTML_Template.html"

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
    [xml]$script:XAML = Get-Content -Path $script:XamlFile -Raw
    Write-Log "Loaded XAML template from: $script:XamlFile"
    
    $script:HTMLTemplate = Get-Content -Path $script:HtmlTemplateFile -Raw
    Write-Log "Loaded HTML template from: $script:HtmlTemplateFile"
}
catch {
    Write-Error "Error loading resource files: $_"
    return
}

# Add this function before New-HTMLReport
function Format-OUPath {
    param(
        [string]$OUPath,
        [int]$GroupCount = 0
    )
    
    if ([string]::IsNullOrEmpty($OUPath)) { return '' }
    
    # Split the DN into parts and extract useful information
    $parts = $OUPath -split ',' | Where-Object { $_ -match '^(OU|DC)=' }
    
    # Get OU parts safely with null checks
    $ouParts = @($parts | Where-Object { $_ -match '^OU=' })
    if ($ouParts.Count -eq 0) { return '' }
    
    # Extract current OU (child) and immediate parent
    $currentOU = ($ouParts[0] -replace '^OU=','').Trim()
    $parentOU = if ($ouParts.Count -gt 1) {
        ($ouParts[1] -replace '^OU=','').Trim()
    } else { $null }
    
    # Build the display string
    if ($parentOU) {
        "<div class='ou-path'>$parentOU / $currentOU</div>"
    } else {
        "<div class='ou-path'>$currentOU</div>"
    }
}

# ... rest of the existing code ...

# In the New-HTMLReport function, replace the HTML string with template loading
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
            } else { $null }
            
            # Calculate disabled percentage
            $disabledPercentage = if ($stats.TotalMembers -gt 0) {
                [math]::Round(($stats.DisabledMembers / $stats.TotalMembers) * 100, 1)
            } else { 0 }
            
            @{
                OU = $currentOU
                ParentOU = $parentOU
                FullDN = $fullDN
                Count = $stats.GroupCount
                EnabledMembers = $stats.EnabledMembers
                DisabledMembers = $stats.DisabledMembers
                DisabledPercentage = $disabledPercentage
                TotalNested = $stats.NestedGroupCount
                MaxDepth = $stats.MaxNestingDepth
            }
        } | Sort-Object { $_.Count } -Descending

        # Generate OU statistics table rows
        $ouStatsRows = $ouStats | ForEach-Object {
            $disabledClass = if ($_.DisabledPercentage -gt 20) { 'warning-text' } elseif ($_.DisabledPercentage -gt 40) { 'critical-text' } else { '' }
            $nestingClass = if ($_.MaxDepth -gt 5) { 'warning-text' } elseif ($_.MaxDepth -gt 10) { 'critical-text' } else { '' }
            @"
            <tr>
                <td>
                    <div class="ou-container">
                        <div class="ou-name">$($_.OU)</div>
                        $(if ($_.ParentOU) {
                            @"
                            <button class="ou-info-button" onclick="toggleOUInfo(this)">Parent OU Info</button>
                            <div class="ou-info">
                                <div class="ou-parent">Parent: $($_.ParentOU)</div>
                                <div class="ou-full-dn">Full DN: $($_.FullDN)</div>
                            </div>
"@
                        })
                    </div>
                </td>
                <td class="members-cell">$($_.Count)</td>
                <td class="members-cell">$($_.EnabledMembers) / $($_.DisabledMembers)</td>
                <td class="members-cell">$($_.EnabledMembers + $_.DisabledMembers)</td>
                <td class="members-cell $disabledClass">$($_.DisabledPercentage)%</td>
                <td class="nesting-cell">$($_.TotalNested)</td>
                <td class="nesting-cell $nestingClass">$($_.MaxDepth)</td>
            </tr>
"@
        }

        # Generate group rows
        $groupRows = $Groups | Sort-Object -Unique Name | Sort-Object @{Expression={$_.TotalMembers}; Descending=$true}, Name | ForEach-Object {
            $healthBadge = if ($_.HealthScore -le 50) {
                'badge-critical'
            } elseif ($_.HealthScore -le 80) {
                'badge-warning'
            } else {
                'badge-success'
            }
            
            @"
            <tr>
                <td>
                    <strong>$($_.Name)</strong>
                    <div class="group-details">
                        <span>$($_.SamAccountName)</span>
                        <span>$($_.OU)</span>
                    </div>
                </td>
                <td>
                    <span class="badge $healthBadge health-score">$($_.HealthScore)</span>
                </td>
                <td>
                    <div class="member-stats">
                        <span class="member-stat">
                            <svg viewBox="0 0 24 24" width="16" height="16" style="margin-right: 4px;">
                                <path fill="currentColor" d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/>
                            </svg>
                            Total: $($_.TotalMembers)
                        </span>
                        <span class="member-stat">
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5z"/>
                            </svg>
                            Users: $($_.UserMembers)
                        </span>
                        <span class="member-stat">
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M12 7V3H2v18h20V7H12zM6 19H4v-2h2v2zm0-4H4v-2h2v2zm0-4H4V9h2v2zm0-4H4V5h2v2zm4 12H8v-2h2v2zm0-4H8v-2h2v2zm0-4H8V9h2v2zm0-4H8V5h2v2zm10 12h-8v-2h2v-2h-2v-2h2v-2h-2V9h8v10zm-2-8h-2v2h2v-2z"/>
                            </svg>
                            Groups: $($_.GroupMembers)
                        </span>
                        <span class="member-stat">
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M21 2H3c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h7l-2 3v1h8v-1l-2-3h7c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z"/>
                            </svg>
                            Computers: $($_.ComputerMembers)
                        </span>
                    </div>
                </td>
                <td>
                    <div class="description">$($_.Description)</div>
                    $(if ($_.Info) {"<div class='notes'>$($_.Info)</div>"})
                    $(if ($_.HealthIssues) {
                        "<div class='issues-list'>" + 
                        ($_.HealthIssues | ForEach-Object { 
                            "<div class='issue-item'>$([System.Web.HttpUtility]::HtmlEncode($_))</div>" 
                        }) -join "`n" +
                        "</div>"
                    })
                </td>
                <td>
                    $(if ($_.Manager) {
                        @"
                        <div class="manager-info">
                            <span class="manager-name">$([System.Web.HttpUtility]::HtmlEncode($_.Manager.DisplayName))</span>
                            <span class="manager-title">$([System.Web.HttpUtility]::HtmlEncode($_.Manager.Title))</span>
                            <span class="manager-upn">$([System.Web.HttpUtility]::HtmlEncode($_.Manager.UPN))</span>
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
                            Category: $($_.Category)
                        </div>
                        <div>
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8z"/>
                            </svg>
                            Scope: $($_.Scope)
                        </div>
                        <div>
                            <svg viewBox="0 0 24 24" width="14" height="14" style="margin-right: 4px;">
                                <path fill="currentColor" d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.11 0-1.99.9-1.99 2L3 19c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z"/>
                            </svg>
                            Created: $($_.Created.ToString('yyyy-MM-dd'))
                        </div>
                        $(if ($_.HasNestedGroups -or $_.NestedInGroupCount -gt 0) {
                            "<div class='nested-warning'>" +
                                $(if ($_.NestedInGroupCount -gt 0) { 
                                    "Member of $($_.NestedInGroupCount) parent groups"
                                }) +
                                $(if ($_.HasNestedGroups -and $_.NestedInGroupCount -gt 0) { "<br/>" }) +
                                $(if ($_.HasNestedGroups) { 
                                    "Contains nested group members"
                                }) +
                            "</div>" +
                            "<div class='nested-groups-container'>" +
                                "<button class='expand-button' onclick='toggleNestedGroups(this)'>Nested Group Info</button>" +
                                "<div class='nested-groups-details'>" +
                                    $(if ($_.NestedInGroupCount -gt 0) {
                                        "<h4>Parent Groups</h4>" +
                                        "<ul class='nested-groups-list'>" +
                                        ($_.ParentGroups | ForEach-Object { "<li>$_</li>" } | Out-String) +
                                        "</ul>"
                                    }) +
                                    $(if ($_.HasNestedGroups) {
                                        $(if ($_.NestedInGroupCount -gt 0) { "<br/>" }) +
                                        "<h4>Nested Groups</h4>" +
                                        "<ul class='nested-groups-list'>" +
                                        ($_.NestedGroups | ForEach-Object { "<li>$_</li>" } | Out-String) +
                                        "</ul>"
                                    }) +
                                "</div>" +
                            "</div>"
                        })
                    </div>
                </td>
            </tr>
"@
        }

        # Load and format HTML template
        $HTML = $script:HTMLTemplate
        
        # Replace placeholders with values
        $replacements = @{
            '{{TOTAL_GROUPS}}' = [int]$totalGroups
            '{{AVG_HEALTH}}' = [Math]::Round([double]($avgHealth ?? 0), 1)
            '{{CRITICAL_GROUPS}}' = [int]$criticalGroups
            '{{WARNING_GROUPS}}' = [int]$warningGroups
            '{{HEALTHY_GROUPS}}' = [int]$healthyGroups
            '{{EMPTY_GROUPS}}' = [int]$emptyGroups
            '{{NO_MANAGER}}' = [int]$noManager
            '{{NO_DESCRIPTION}}' = [int]$noDescription
            '{{NESTED_GROUPS}}' = [int]$nestedGroups
            '{{OU_STATS_ROWS}}' = $ouStatsRows -join "`n"
            '{{GROUP_ROWS}}' = $groupRows -join "`n"
        }
        
        foreach ($key in $replacements.Keys) {
            $HTML = $HTML.Replace($key, $replacements[$key])
        }
        
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
                # Removed unused variable assignment
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

# ... rest of the code ...

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