Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check for administrative privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    [System.Windows.Forms.MessageBox]::Show("This script requires administrative privileges. Please run PowerShell as an administrator.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Active Directory and System Management Tool"
$form.Size = New-Object System.Drawing.Size(1000, 800)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 800)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

# Create a ToolTip object
$toolTip = New-Object System.Windows.Forms.ToolTip

# Create panels for better organization
$inputPanel = New-Object System.Windows.Forms.Panel
$inputPanel.Location = New-Object System.Drawing.Point(10, 10)
$inputPanel.Size = New-Object System.Drawing.Size(330, 750)
$inputPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$inputPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($inputPanel)

$outputPanel = New-Object System.Windows.Forms.Panel
$outputPanel.Location = New-Object System.Drawing.Point(350, 10)
$outputPanel.Size = New-Object System.Drawing.Size(640, 750)
$outputPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$outputPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($outputPanel)

# Create GroupBox for AD User and Group
$adGroupBox = New-Object System.Windows.Forms.GroupBox
$adGroupBox.Text = "AD User and Group"
$adGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$adGroupBox.Size = New-Object System.Drawing.Size(320, 190)
$inputPanel.Controls.Add($adGroupBox)

# Create GroupBox for Device Optimization
$deviceGroupBox = New-Object System.Windows.Forms.GroupBox
$deviceGroupBox.Text = "Device Optimization"
$deviceGroupBox.Location = New-Object System.Drawing.Point(10, 210)
$deviceGroupBox.Size = New-Object System.Drawing.Size(320, 220)
$inputPanel.Controls.Add($deviceGroupBox)

# Create GroupBox for Troubleshooting
$troubleGroupBox = New-Object System.Windows.Forms.GroupBox
$troubleGroupBox.Text = "Troubleshooting"
$troubleGroupBox.Location = New-Object System.Drawing.Point(10, 440)
$troubleGroupBox.Size = New-Object System.Drawing.Size(320, 210)
$inputPanel.Controls.Add($troubleGroupBox)

# --- AD User and Group Controls ---
$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Text = "Username:"
$userLabel.Location = New-Object System.Drawing.Point(10, 20)
$userLabel.AutoSize = $true
$adGroupBox.Controls.Add($userLabel)

$userTextBox = New-Object System.Windows.Forms.TextBox
$userTextBox.Location = New-Object System.Drawing.Point(10, 40)
$userTextBox.Size = New-Object System.Drawing.Size(190, 20)
$userTextBox.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) {
        $userButton.PerformClick()
    }
})
$adGroupBox.Controls.Add($userTextBox)

$userButton = New-Object System.Windows.Forms.Button
$userButton.Text = "Get User Details"
$userButton.Location = New-Object System.Drawing.Point(210, 40)
$userButton.Size = New-Object System.Drawing.Size(100, 23)
$userButton.Add_Click({
    $outputTextBox.Clear()
    $username = $userTextBox.Text.Trim()
    if (-not $username) {
        $outputTextBox.Text = "Please enter a username."
        return
    }
    try {
        $output = net user $username /domain 2>&1
        if ($LASTEXITCODE -eq 0 -and $output -match "User name") {
            $outputTextBox.Lines = $output
        } else {
            $outputTextBox.Text = "No user found with the exact username '$username' or error occurred."
        }
    } catch {
        $outputTextBox.Text = "Error retrieving user: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($userButton, "Get detailed information for exact username match")
$adGroupBox.Controls.Add($userButton)

$searchUserButton = New-Object System.Windows.Forms.Button
$searchUserButton.Text = "Search Users"
$searchUserButton.Location = New-Object System.Drawing.Point(210, 70)
$searchUserButton.Size = New-Object System.Drawing.Size(90, 23)
$searchUserButton.Add_Click({
    $outputTextBox.Clear()
    $username = $userTextBox.Text.Trim()
    if (-not $username) {
        $outputTextBox.Text = "Please enter a username (partial or full)."
        return
    }
    try {
        $output = net user /domain 2>&1
        if ($LASTEXITCODE -eq 0) {
            $users = $output | Where-Object { $_ -match $username }
            if ($users) {
                $outputTextBox.Lines = $users
            } else {
                $outputTextBox.Text = "No users found matching '$username'."
            }
        } else {
            $outputTextBox.Text = "Error retrieving users."
        }
    } catch {
        $outputTextBox.Text = "Error searching users: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($searchUserButton, "Search for users with partial name match")
$adGroupBox.Controls.Add($searchUserButton)

$groupLabel = New-Object System.Windows.Forms.Label
$groupLabel.Text = "Group Name:"
$groupLabel.Location = New-Object System.Drawing.Point(10, 100)
$groupLabel.AutoSize = $true
$adGroupBox.Controls.Add($groupLabel)

$groupTextBox = New-Object System.Windows.Forms.TextBox
$groupTextBox.Location = New-Object System.Drawing.Point(10, 120)
$groupTextBox.Size = New-Object System.Drawing.Size(190, 20)
$groupTextBox.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) {
        $groupButton.PerformClick()
    }
})
$adGroupBox.Controls.Add($groupTextBox)

$groupButton = New-Object System.Windows.Forms.Button
$groupButton.Text = "Get Group Details"
$groupButton.Location = New-Object System.Drawing.Point(210, 120)
$groupButton.Size = New-Object System.Drawing.Size(100, 23)
$groupButton.Add_Click({
    $outputTextBox.Clear()
    $groupname = $groupTextBox.Text.Trim()
    if (-not $groupname) {
        $outputTextBox.Text = "Please enter a group name."
        return
    }
    try {
        $output = net group $groupname /domain 2>&1
        if ($LASTEXITCODE -eq 0 -and $output -match "Group name") {
            $outputTextBox.Lines = $output
        } else {
            $outputTextBox.Text = "No group found with the exact name '$groupname' or error occurred."
        }
    } catch {
        $outputTextBox.Text = "Error retrieving group: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($groupButton, "Get detailed information for exact group name match")
$adGroupBox.Controls.Add($groupButton)

$searchGroupButton = New-Object System.Windows.Forms.Button
$searchGroupButton.Text = "Search Groups"
$searchGroupButton.Location = New-Object System.Drawing.Point(210, 150)
$searchGroupButton.Size = New-Object System.Drawing.Size(90, 23)
$searchGroupButton.Add_Click({
    $outputTextBox.Clear()
    $groupname = $groupTextBox.Text.Trim()
    if (-not $groupname) {
        $outputTextBox.Text = "Please enter a group name (partial or full)."
        return
    }
    try {
        $output = net group /domain 2>&1
        if ($LASTEXITCODE -eq 0) {
            $groups = $output | Where-Object { $_ -match $groupname }
            if ($groups) {
                $outputTextBox.Lines = $groups
            } else {
                $outputTextBox.Text = "No groups found matching '$groupname'."
            }
        } else {
            $outputTextBox.Text = "Error retrieving groups."
        }
    } catch {
        $outputTextBox.Text = "Error searching groups: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($searchGroupButton, "Search for groups with partial name match")
$adGroupBox.Controls.Add($searchGroupButton)

$getCompInfoButton = New-Object System.Windows.Forms.Button
$getCompInfoButton.Text = "Get Computer Info"
$getCompInfoButton.Location = New-Object System.Drawing.Point(120, 670)
$getCompInfoButton.Size = New-Object System.Drawing.Size(130, 23)
$getCompInfoButton.Add_Click({
    $outputTextBox.Clear()
    try {
        $outputTextBox.Text = "Retrieving computer information...`n`n"
        $computerInfo = Get-ComputerInfo -ErrorAction Stop
        $outputTextBox.AppendText("Brand: $($computerInfo.CsManufacturer)`n")
	$outputTextBox.AppendText("Model: $($computerInfo.CsModel)`n")
	$outputTextBox.AppendText("Serial Number: $($computerInfo.BiosSeralNumber)`n")
	$outputTextBox.AppendText("Computer Name: $($computerInfo.CsName)`n")
        $outputTextBox.AppendText("Domain: $($computerInfo.CsDomain)`n")
	$outputTextBox.AppendText("Logon Server: $($computerInfo.LogonServer)`n")
	$outputTextBox.AppendText("Time Zone: $($computerInfo.TimeZone)`n")
	$outputTextBox.AppendText("Current User: $($computerInfo.CsUserName)`n")
	$outputTextBox.AppendText("Management Type: $($computerInfo.WindowsRegisteredOwner)`n")
        $outputTextBox.AppendText("Windows Product Name: $($computerInfo.WindowsProductName)`n")
	$outputTextBox.AppendText("Current OS: $($computerInfo.OsName)`n")
        $outputTextBox.AppendText("OS Version: $($computerInfo.OsVersion)`n")
        $outputTextBox.AppendText("OS Build: $($computerInfo.OsBuildNumber)`n")
        $outputTextBox.AppendText("Total Physical Memory: $([math]::Round($computerInfo.CsTotalPhysicalMemory / 1GB, 2)) GB`n")
        $outputTextBox.AppendText("Processor: $($computerInfo.CsProcessors[0].Name)`n")
        $outputTextBox.AppendText("BIOS Version: $($computerInfo.BiosBIOSVersion -join ', ')`n")
        $outputTextBox.AppendText("Last Boot Time: $($computerInfo.OsLastBootUpTime)`n")
        $outputTextBox.AppendText("System Type: $($computerInfo.CsSystemType)`n")
        $outputTextBox.AppendText("`nComputer information retrieved successfully.")
    } catch {
        $outputTextBox.Text = "Error retrieving computer information: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($getCompInfoButton, "Display detailed computer information including Device, User, OS, hardware, and domain details")
$inputPanel.Controls.Add($getCompInfoButton)

# --- Device Optimization Controls ---
$optimizePowerButton = New-Object System.Windows.Forms.Button
$optimizePowerButton.Text = "Optimize Power"
$optimizePowerButton.Location = New-Object System.Drawing.Point(10, 20)
$optimizePowerButton.Size = New-Object System.Drawing.Size(100, 23)
$optimizePowerButton.Add_Click({
    $outputTextBox.Clear()
    try {
        $hibernationStatus = powercfg /a 2>&1
        $hibernationEnabled = $hibernationStatus -notmatch "Hibernation has not been enabled" -and $hibernationStatus -notmatch "The following sleep states are not available"
        
        if ($hibernationEnabled) {
            $output = powercfg -h off 2>&1
            if ($LASTEXITCODE -ne 0) {
                $outputTextBox.AppendText("Failed to disable hibernation. Error: $output`n")
            } else {
                $outputTextBox.AppendText("Hibernation disabled successfully.`n")
            }
        } else {
            $outputTextBox.AppendText("Hibernation is already disabled or not supported on this system.`n")
        }

        $output = powercfg -change -monitor-timeout-dc 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set display timeout (On battery). Error: $output`n") }
        else { $outputTextBox.AppendText("Display timeout set to Never (On battery).`n") }
        
        $output = powercfg -change -monitor-timeout-ac 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set display timeout (Plugged in). Error: $output`n") }
        else { $outputTextBox.AppendText("Display timeout set to Never (Plugged in).`n") }

        $output = powercfg -change -standby-timeout-dc 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set sleep timeout (On battery). Error: $output`n") }
        else { $outputTextBox.AppendText("Sleep timeout set to Never (On battery).`n") }
        
        $output = powercfg -change -standby-timeout-ac 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set sleep timeout (Plugged in). Error: $output`n") }
        else { $outputTextBox.AppendText("Sleep timeout set to Never (Plugged in).`n") }

        $output = powercfg -change -disk-timeout-dc 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set disk timeout (On battery). Error: $output`n") }
        else { $outputTextBox.AppendText("Hard disk timeout set to Never (On battery).`n") }
        
        $output = powercfg -change -disk-timeout-ac 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set disk timeout (Plugged in). Error: $output`n") }
        else { $outputTextBox.AppendText("Hard disk timeout set to Never (Plugged in).`n") }

        $output = powercfg -setdcvalueindex SCHEME_CURRENT SUB_BUTTONS PBUTTONACTION 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set power button action (On battery). Error: $output`n") }
        else { $outputTextBox.AppendText("Power button action set to Do nothing (On battery).`n") }
        
        $output = powercfg -setacvalueindex SCHEME_CURRENT SUB_BUTTONS PBUTTONACTION 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set power button action (Plugged in). Error: $output`n") }
        else { $outputTextBox.AppendText("Power button action set to Do nothing (Plugged in).`n") }

        $output = powercfg -setdcvalueindex SCHEME_CURRENT SUB_BUTTONS SBUTTONACTION 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set sleep button action (On battery). Error: $output`n") }
        else { $outputTextBox.AppendText("Sleep button action set to Do nothing (On battery).`n") }
        
        $output = powercfg -setacvalueindex SCHEME_CURRENT SUB_BUTTONS SBUTTONACTION 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set sleep button action (Plugged in). Error: $output`n") }
        else { $outputTextBox.AppendText("Sleep button action set to Do nothing (Plugged in).`n") }

        $output = powercfg -setdcvalueindex SCHEME_CURRENT SUB_BUTTONS LIDACTION 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set lid close action (On battery). Error: $output`n") }
        else { $outputTextBox.AppendText("Lid close action set to Do nothing (On battery).`n") }
        
        $output = powercfg -setacvalueindex SCHEME_CURRENT SUB_BUTTONS LIDACTION 0 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to set lid close action (Plugged in). Error: $output`n") }
        else { $outputTextBox.AppendText("Lid close action set to Do nothing (Plugged in).`n") }

        $output = powercfg -setactive SCHEME_CURRENT 2>&1
        if ($LASTEXITCODE -ne 0) { $outputTextBox.AppendText("Failed to apply power settings. Error: $output`n") }
        else { $outputTextBox.AppendText("Power settings applied successfully.`n") }

        $outputTextBox.AppendText("Power optimization completed. Check above for any errors.")
    } catch {
        $outputTextBox.AppendText("Unexpected error during power optimization: $($_.Exception.Message)`n")
        $outputTextBox.AppendText("Power optimization completed with errors. Check above for details.")
    }
})
$toolTip.SetToolTip($optimizePowerButton, "Optimize power settings: disable hibernation, set timeouts to Never, set buttons to Do nothing")
$deviceGroupBox.Controls.Add($optimizePowerButton)

$optimizePCButton = New-Object System.Windows.Forms.Button
$optimizePCButton.Text = "Optimize PC"
$optimizePCButton.Location = New-Object System.Drawing.Point(120, 20)
$optimizePCButton.Size = New-Object System.Drawing.Size(100, 23)
$optimizePCButton.Add_Click({
    $outputTextBox.Clear()
    try {
        $totalFreedSpace = 0

        function Get-FolderSize {
            param ($Path)
            try {
                if (Test-Path $Path) {
                    $size = (Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                    return [math]::Round($size / 1MB, 2)
                }
                return 0
            } catch {
                return 0
            }
        }

        $tempPath = [System.IO.Path]::GetTempPath()
        $tempSizeBefore = Get-FolderSize -Path $tempPath
        try {
            Get-ChildItem -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match "Access is denied") {
                        $outputTextBox.AppendText("Skipped Temporary File '$($_.FullName)' due to access denied.`n")
                    } else {
                        $outputTextBox.AppendText("Error deleting Temporary File '$($_.FullName)': $($_.Exception.Message)`n")
                    }
                }
            }
            $tempSizeAfter = Get-FolderSize -Path $tempPath
            $freed = $tempSizeBefore - $tempSizeAfter
            $totalFreedSpace += $freed
            $outputTextBox.AppendText("Temporary Files cleaned. Freed: $freed MB`n")
        } catch {
            $outputTextBox.AppendText("Error cleaning Temporary Files: $($_.Exception.Message)`n")
        }

        try {
            $recycleSizeBefore = Get-FolderSize -Path "C:\$Recycle.Bin"
            if (Get-ChildItem -Path "C:\$Recycle.Bin" -Recurse -Force -ErrorAction SilentlyContinue) {
                $output = Clear-RecycleBin -DriveLetter C -Force -ErrorAction Stop 2>&1
                $recycleSizeAfter = Get-FolderSize -Path "C:\$Recycle.Bin"
                $freed = $recycleSizeBefore - $recycleSizeAfter
                $totalFreedSpace += $freed
                $outputTextBox.AppendText("Recycle Bin emptied. Freed: $freed MB`n")
            } else {
                $outputTextBox.AppendText("Recycle Bin is already empty.`n")
            }
        } catch {
            $outputTextBox.AppendText("Error emptying Recycle Bin: $($_.Exception.Message)`n")
        }

        $prefetchPath = "C:\Windows\Prefetch"
        $prefetchSizeBefore = Get-FolderSize -Path $prefetchPath
        try {
            Get-ChildItem -Path $prefetchPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match "Access is denied") {
                        $outputTextBox.AppendText("Skipped Prefetch File '$($_.FullName)' due to access denied.`n")
                    } else {
                        $outputTextBox.AppendText("Error deleting Prefetch File '$($_.FullName)': $($_.Exception.Message)`n")
                    }
                }
            }
            $prefetchSizeAfter = Get-FolderSize -Path $prefetchPath
            $freed = $prefetchSizeBefore - $prefetchSizeAfter
            $totalFreedSpace += $freed
            $outputTextBox.AppendText("Prefetch Files cleaned. Freed: $freed MB`n")
        } catch {
            $outputTextBox.AppendText("Error cleaning Prefetch Files: $($_.Exception.Message)`n")
        }

        $updatePath = "C:\Windows\SoftwareDistribution\Download"
        $updateSizeBefore = Get-FolderSize -Path $updatePath
        try {
            Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path $updatePath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match "Access is denied") {
                        $outputTextBox.AppendText("Skipped Windows Update File '$($_.FullName)' due to access denied.`n")
                    } else {
                        $outputTextBox.AppendText("Error deleting Windows Update File '$($_.FullName)': $($_.Exception.Message)`n")
                    }
                }
            }
            Start-Service -Name wuauserv -ErrorAction SilentlyContinue
            $updateSizeAfter = Get-FolderSize -Path $updatePath
            $freed = $updateSizeBefore - $updateSizeAfter
            $totalFreedSpace += $freed
            $outputTextBox.AppendText("Old Windows Update Files cleaned. Freed: $freed MB`n")
        } catch {
            $outputTextBox.AppendText("Error cleaning Windows Update Files: $($_.Exception.Message)`n")
        }

        $browserPaths = @(
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
            "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*\cache",
            "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*\cache2"
        )
        foreach ($path in $browserPaths) {
            try {
                $browserSizeBefore = Get-FolderSize -Path $path
                Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    try {
                        Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                    } catch {
                        if ($_.Exception.Message -match "Access is denied") {
                            $outputTextBox.AppendText("Skipped Browser Cache File '$($_.FullName)' due to access denied.`n")
                        } else {
                            $outputTextBox.AppendText("Error deleting Browser Cache File '$($_.FullName)': $($_.Exception.Message)`n")
                        }
                    }
                }
                $browserSizeAfter = Get-FolderSize -Path $path
                $freed = $browserSizeBefore - $browserSizeAfter
                $totalFreedSpace += $freed
                $outputTextBox.AppendText("Browser Cache ($path) cleaned. Freed: $freed MB`n")
            } catch {
                $outputTextBox.AppendText("Error cleaning Browser Cache ($path): $($_.Exception.Message)`n")
            }
        }

        $errorReportPath = "$env:PROGRAMDATA\Microsoft\Windows\WER\ReportArchive"
        $errorReportSizeBefore = Get-FolderSize -Path $errorReportPath
        try {
            Get-ChildItem -Path $errorReportPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match "Access is denied") {
                        $outputTextBox.AppendText("Skipped Error Reporting File '$($_.FullName)' due to access denied.`n")
                    } else {
                        $outputTextBox.AppendText("Error deleting Error Reporting File '$($_.FullName)': $($_.Exception.Message)`n")
                    }
                }
            }
            $errorReportSizeAfter = Get-FolderSize -Path $errorReportPath
            $freed = $errorReportSizeBefore - $errorReportSizeAfter
            $totalFreedSpace += $freed
            $outputTextBox.AppendText("Windows Error Reporting Files cleaned. Freed: $freed MB`n")
        } catch {
            $outputTextBox.AppendText("Error cleaning Windows Error Reporting Files: $($_.Exception.Message)`n")
        }

        $memoryDumpPath = "C:\Windows\*.dmp"
        $memoryDumpSizeBefore = Get-FolderSize -Path $memoryDumpPath
        try {
            Get-ChildItem -Path $memoryDumpPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match "Access is denied") {
                        $outputTextBox.AppendText("Skipped Memory Dump File '$($_.FullName)' due to access denied.`n")
                    } else {
                        $outputTextBox.AppendText("Error deleting Memory Dump File '$($_.FullName)': $($_.Exception.Message)`n")
                    }
                }
            }
            $memoryDumpSizeAfter = Get-FolderSize -Path $memoryDumpPath
            $freed = $memoryDumpSizeBefore - $memoryDumpSizeAfter
            $totalFreedSpace += $freed
            $outputTextBox.AppendText("Memory Dump Files cleaned. Freed: $freed MB`n")
        } catch {
            $outputTextBox.AppendText("Error cleaning Memory Dump Files: $($_.Exception.Message)`n")
        }

        try {
            $vssService = Get-Service -Name VSS -ErrorAction SilentlyContinue
            if ($vssService) {
                Start-Service -Name VSS -ErrorAction SilentlyContinue
                $shadows = vssadmin List Shadows 2>&1
                if ($shadows -match "Shadow Copy ID") {
                    $output = vssadmin Delete Shadows /All /Quiet 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        $freed = [math]::Round((Get-ChildItem -Path "C:\System Volume Information" -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum / 1MB, 2)
                        $totalFreedSpace += $freed
                        $outputTextBox.AppendText("Old System Restore Points cleaned. Freed: $freed MB`n")
                    } else {
                        $outputTextBox.AppendText("Error cleaning System Restore Points: $output`n")
                    }
                } else {
                    $outputTextBox.AppendText("No System Restore Points found.`n")
                }
            } else {
                $outputTextBox.AppendText("Volume Shadow Copy service not found. Skipping System Restore Points cleanup.`n")
            }
        } catch {
            $outputTextBox.AppendText("Error cleaning System Restore Points: $($_.Exception.Message)`n")
        }

        $deliveryOptPath = "$env:PROGRAMDATA\Microsoft\Network\Downloader"
        $deliveryOptSizeBefore = Get-FolderSize -Path $deliveryOptPath
        try {
            $doService = Get-Service -Name DoSvc -ErrorAction SilentlyContinue
            if ($doService) {
                Stop-Service -Name DoSvc -Force -ErrorAction SilentlyContinue
                Get-ChildItem -Path $deliveryOptPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    try {
                        Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                    } catch {
                        if ($_.Exception.Message -match "Access is denied") {
                            $outputTextBox.AppendText("Skipped Delivery Optimization File '$($_.FullName)' due to access denied.`n")
                        } else {
                            $outputTextBox.AppendText("Error deleting Delivery Optimization File '$($_.FullName)': $($_.Exception.Message)`n")
                        }
                    }
                }
                Start-Service -Name DoSvc -ErrorAction SilentlyContinue
                $deliveryOptSizeAfter = Get-FolderSize -Path $deliveryOptPath
                $freed = $deliveryOptSizeBefore - $deliveryOptSizeAfter
                $totalFreedSpace += $freed
                $outputTextBox.AppendText("Delivery Optimization Cache cleaned. Freed: $freed MB`n")
            } else {
                $outputTextBox.AppendText("Delivery Optimization service not found. Skipping cache cleanup.`n")
            }
        } catch {
            $outputTextBox.AppendText("Error cleaning Delivery Optimization Cache: $($_.Exception.Message)`n")
        }

        $tempLogsPath = "C:\Windows\Logs"
        $tempLogsSizeBefore = Get-FolderSize -Path $tempLogsPath
        try {
            Get-ChildItem -Path $tempLogsPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -Recurse -ErrorAction Stop
                } catch {
                    if ($_.Exception.Message -match "Access is denied") {
                        $outputTextBox.AppendText("Skipped Windows Temp Log '$($_.FullName)' due to access denied.`n")
                    } else {
                        $outputTextBox.AppendText("Error deleting Windows Temp Log '$($_.FullName)': $($_.Exception.Message)`n")
                    }
                }
            }
            $tempLogsSizeAfter = Get-FolderSize -Path $tempLogsPath
            $freed = $tempLogsSizeBefore - $tempLogsSizeAfter
            $totalFreedSpace += $freed
            $outputTextBox.AppendText("Windows Temp Logs cleaned. Freed: $freed MB`n")
        } catch {
            $outputTextBox.AppendText("Error cleaning Windows Temp Logs: $($_.Exception.Message)`n")
        }

        $outputTextBox.AppendText("`nPC Optimization completed. Total space freed: $totalFreedSpace MB`n")
        $outputTextBox.AppendText("Check above for any errors.")
    } catch {
        $outputTextBox.AppendText("Unexpected error during PC optimization: $($_.Exception.Message)`n")
        $outputTextBox.AppendText("PC optimization completed with errors. Check above for details.")
    }
})
$toolTip.SetToolTip($optimizePCButton, "Clean Temporary Files, Recycle Bin, Prefetch, Windows Update Files, Browser Cache, Error Reports, Memory Dumps, System Restore Points, Delivery Optimization Cache, and Windows Temp Logs")
$deviceGroupBox.Controls.Add($optimizePCButton)

$filePathLabel = New-Object System.Windows.Forms.Label
$filePathLabel.Text = "File Path:"
$filePathLabel.Location = New-Object System.Drawing.Point(10, 50)
$filePathLabel.AutoSize = $true
$deviceGroupBox.Controls.Add($filePathLabel)

$filePathTextBox = New-Object System.Windows.Forms.TextBox
$filePathTextBox.Location = New-Object System.Drawing.Point(10, 70)
$filePathTextBox.Size = New-Object System.Drawing.Size(190, 20)
$filePathTextBox.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) {
        $openLocationButton.PerformClick()
    }
})
$deviceGroupBox.Controls.Add($filePathTextBox)

$openLocationButton = New-Object System.Windows.Forms.Button
$openLocationButton.Text = "Open Location"
$openLocationButton.Location = New-Object System.Drawing.Point(210, 70)
$openLocationButton.Size = New-Object System.Drawing.Size(90, 23)
$openLocationButton.Add_Click({
    $outputTextBox.Clear()
    $filePath = $filePathTextBox.Text.Trim()
    if (-not $filePath) {
        $outputTextBox.Text = "Error: Please enter a file path."
        return
    }
    try {
        if (Test-Path $filePath -PathType Leaf) {
            $parentPath = Split-Path $filePath -Parent
            Start-Process explorer.exe $parentPath
            $outputTextBox.Text = "Opened folder: $parentPath"
        } else {
            $outputTextBox.Text = "Error: File '$filePath' does not exist."
        }
    } catch {
        $outputTextBox.Text = "Error opening file location: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($openLocationButton, "Open the parent folder of the specified file in File Explorer")
$deviceGroupBox.Controls.Add($openLocationButton)

$deleteFileButton = New-Object System.Windows.Forms.Button
$deleteFileButton.Text = "Delete File"
$deleteFileButton.Location = New-Object System.Drawing.Point(210, 100)
$deleteFileButton.Size = New-Object System.Drawing.Size(90, 23)
$deleteFileButton.Add_Click({
    $outputTextBox.Clear()
    $filePath = $filePathTextBox.Text.Trim()
    if (-not $filePath) {
        $outputTextBox.Text = "Error: Please enter a file path."
        return
    }
    try {
        if (Test-Path $filePath -PathType Leaf) {
            $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete '$filePath'?", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($confirm -eq "Yes") {
                Remove-Item -Path $filePath -Force -ErrorAction Stop
                $outputTextBox.Text = "File '$filePath' deleted successfully."
            } else {
                $outputTextBox.Text = "File deletion cancelled."
            }
        } else {
            $outputTextBox.Text = "Error: File '$filePath' does not exist."
        }
    } catch {
        $outputTextBox.Text = "Error deleting file: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($deleteFileButton, "Delete the specified file after confirmation")
$deviceGroupBox.Controls.Add($deleteFileButton)

$findLargeFilesButton = New-Object System.Windows.Forms.Button
$findLargeFilesButton.Text = "Find Largest Files"
$findLargeFilesButton.Location = New-Object System.Drawing.Point(10, 100)
$findLargeFilesButton.Size = New-Object System.Drawing.Size(120, 23)
$findLargeFilesButton.Add_Click({
    $outputTextBox.Clear()
    try {
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select a drive or folder to search for the largest files"
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $path = $folderBrowser.SelectedPath
            if (-not (Test-Path $path)) {
                $outputTextBox.Text = "Error: Selected path '$path' is invalid."
                return
            }
            $outputTextBox.Text = "Searching for largest files in '$path'...`n`n"
            $files = Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue |
                Sort-Object Length -Descending |
                Select-Object -First 10 |
                ForEach-Object {
                    [PSCustomObject]@{
                        FileName = $_.Name
                        FullPath = $_.FullName
                        FileType = $_.Extension
                        SizeMB   = [math]::Round($_.Length / 1MB, 2)
                    }
                }
            if ($files) {
                $index = 1
                foreach ($file in $files) {
                    $outputTextBox.AppendText("File $index`n")
                    $outputTextBox.AppendText("File Name: $($file.FileName)`n")
                    $outputTextBox.AppendText("Full Path: $($file.FullPath)`n")
                    $outputTextBox.AppendText("File Type: $($file.FileType)`n")
                    $outputTextBox.AppendText("Size: $($file.SizeMB) MB`n`n")
                    $index++
                }
                $outputTextBox.AppendText("Top 10 largest files listed above.")
            } else {
                $outputTextBox.Text += "No files found in the specified location."
            }
        } else {
            $outputTextBox.Text = "Search cancelled."
        }
    } catch {
        $outputTextBox.Text = "Error finding largest files: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($findLargeFilesButton, "List the 10 largest files in a selected drive or folder")
$deviceGroupBox.Controls.Add($findLargeFilesButton)

$appIdLabel = New-Object System.Windows.Forms.Label
$appIdLabel.Text = "App ID:"
$appIdLabel.Location = New-Object System.Drawing.Point(10, 130)
$appIdLabel.AutoSize = $true
$deviceGroupBox.Controls.Add($appIdLabel)

$appIdTextBox = New-Object System.Windows.Forms.TextBox
$appIdTextBox.Location = New-Object System.Drawing.Point(10, 150)
$appIdTextBox.Size = New-Object System.Drawing.Size(190, 20)
$appIdTextBox.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) {
        $updateIdButton.PerformClick()
    }
})
$deviceGroupBox.Controls.Add($appIdTextBox)

$updateIdButton = New-Object System.Windows.Forms.Button
$updateIdButton.Text = "Update ID"
$updateIdButton.Location = New-Object System.Drawing.Point(210, 150)
$updateIdButton.Size = New-Object System.Drawing.Size(90, 23)
$updateIdButton.Add_Click({
    $outputTextBox.Clear()
    $appId = $appIdTextBox.Text.Trim()
    if (-not $appId) {
        $outputTextBox.Text = "Error: Please enter an application ID."
        return
    }
    try {
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
            $outputTextBox.Text = "Error: Winget is not installed on this system. Please install it to use this feature."
            return
        }
        $listCheck = winget list --id $appId --accept-source-agreements --disable-interactivity 2>&1
        if ($listCheck -notmatch [regex]::Escape($appId)) {
            $outputTextBox.Text = "Error: Application '$appId' is not installed on this system."
            return
        }
        $upgradeList = winget upgrade --accept-source-agreements --disable-interactivity 2>&1
        if ($upgradeList -imatch [regex]::Escape($appId)) {
            $output = winget upgrade --id "$appId" --accept-source-agreements --disable-interactivity --force 2>&1
            if ($LASTEXITCODE -eq 0) {
                $outputTextBox.Lines = $output
                $outputTextBox.AppendText("`nSuccessfully updated application '$appId'.")
            } else {
                $outputTextBox.Text = "Failed to update application '$appId'. Error: $output"
            }
        } else {
            $outputTextBox.Text = "Error: No updates available for application '$appId' at this time. Raw winget upgrade output for debugging:`n$upgradeList"
        }
    } catch {
        $outputTextBox.Text = "Error updating application '$appId': $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($updateIdButton, "Update the specified application using winget with force option")
$deviceGroupBox.Controls.Add($updateIdButton)

$appUpdateButton = New-Object System.Windows.Forms.Button
$appUpdateButton.Text = "Application Update"
$appUpdateButton.Location = New-Object System.Drawing.Point(10, 180)
$appUpdateButton.Size = New-Object System.Drawing.Size(120, 23)
$appUpdateButton.Add_Click({
    $outputTextBox.Clear()
    try {
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
            $outputTextBox.Text = "Error: Winget is not installed on this system. Please install it from Microsoft Store or https://github.com/microsoft/winget-cli."
            return
        }
        $wingetVersion = winget --version 2>&1
        if ($LASTEXITCODE -ne 0) {
            $outputTextBox.Text = "Error: Unable to verify winget version. Output: $wingetVersion"
            return
        }
        $outputTextBox.Text = "Running winget upgrade (this may take a moment)...`n"
        $regionCode = [System.Globalization.RegionInfo]::CurrentRegion.TwoLetterISORegionName
        $output = winget upgrade --accept-source-agreements --disable-interactivity 2>&1
        if ($LASTEXITCODE -eq 0) {
            if ($output -match "No applicable update found") {
                $outputTextBox.Text = "No application updates found."
            } else {
                $outputTextBox.Lines = $output
                $outputTextBox.AppendText("`nUse the App ID text box and 'Update ID' button to update a specific application.")
            }
        } else {
            $outputTextBox.Text = "Error running winget upgrade: $output`nTry running 'winget source update' manually in an elevated PowerShell."
        }
    } catch {
        $outputTextBox.Text = "Error running winget upgrade: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($appUpdateButton, "List applications with available updates using winget")
$deviceGroupBox.Controls.Add($appUpdateButton)

# --- Troubleshooting Controls ---
$serviceLabel = New-Object System.Windows.Forms.Label
$serviceLabel.Text = "Service Name:"
$serviceLabel.Location = New-Object System.Drawing.Point(10, 20)
$serviceLabel.AutoSize = $true
$troubleGroupBox.Controls.Add($serviceLabel)

$serviceTextBox = New-Object System.Windows.Forms.TextBox
$serviceTextBox.Location = New-Object System.Drawing.Point(10, 40)
$serviceTextBox.Size = New-Object System.Drawing.Size(190, 20)
$serviceTextBox.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) {
        $netStartButton.PerformClick()
    }
})
$troubleGroupBox.Controls.Add($serviceTextBox)

$netStopButton = New-Object System.Windows.Forms.Button
$netStopButton.Text = "Net Stop"
$netStopButton.Location = New-Object System.Drawing.Point(220, 70)
$netStopButton.Size = New-Object System.Drawing.Size(90, 23)
$netStopButton.Add_Click({
    $outputTextBox.Clear()
    $serviceName = $serviceTextBox.Text.Trim()
    if (-not $serviceName) {
        $outputTextBox.Text = "Error: Please enter a service name."
        return
    }
    try {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if (-not $service) {
            $outputTextBox.Text = "Error: Service '$serviceName' does not exist."
            return
        }
        if ($service.Status -eq "Stopped") {
            $outputTextBox.Text = "Service '$serviceName' is already stopped."
            return
        }
        $output = net stop "`"$serviceName`"" 2>&1
        if ($LASTEXITCODE -eq 0) {
            $outputTextBox.Lines = $output
            $outputTextBox.AppendText("`nCurrent Status: " + (Get-Service -Name $serviceName).Status)
        } else {
            $outputTextBox.Text = "Failed to stop service '$serviceName'. Error: $output"
        }
    } catch {
        $outputTextBox.Text = "Error stopping service: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($netStopButton, "Stop the specified service")
$troubleGroupBox.Controls.Add($netStopButton)

$servicesButton = New-Object System.Windows.Forms.Button
$servicesButton.Text = "Get Services"
$servicesButton.Location = New-Object System.Drawing.Point(10, 70)
$servicesButton.Size = New-Object System.Drawing.Size(100, 23)
$servicesButton.Add_Click({
    $outputTextBox.Clear()
    try {
        $services = Get-Service | Select-Object Name, Status, DisplayName | Sort-Object Name
        $output = $services | Format-Table -AutoSize | Out-String
        $outputTextBox.Text = $output
    } catch {
        $outputTextBox.Text = "Error retrieving services: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($servicesButton, "Display list of services with name, status, and display name")
$troubleGroupBox.Controls.Add($servicesButton)

$netStartButton = New-Object System.Windows.Forms.Button
$netStartButton.Text = "Net Start"
$netStartButton.Location = New-Object System.Drawing.Point(120, 70)
$netStartButton.Size = New-Object System.Drawing.Size(90, 23)
$netStartButton.Add_Click({
    $outputTextBox.Clear()
    $serviceName = $serviceTextBox.Text.Trim()
    if (-not $serviceName) {
        $outputTextBox.Text = "Error: Please enter a service name."
        return
    }
    try {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if (-not $service) {
            $outputTextBox.Text = "Error: Service '$serviceName' does not exist."
            return
        }
        if ($service.Status -eq "Running") {
            $outputTextBox.Text = "Service '$serviceName' is already running."
            return
        }
        $output = net start "`"$serviceName`"" 2>&1
        if ($LASTEXITCODE -eq 0) {
            $outputTextBox.Lines = $output
            $outputTextBox.AppendText("`nCurrent Status: " + (Get-Service -Name $serviceName).Status)
        } else {
            $outputTextBox.Text = "Failed to start service '$serviceName'. Error: $output"
        }
    } catch {
        $outputTextBox.Text = "Error starting service: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($netStartButton, "Start the specified service")
$troubleGroupBox.Controls.Add($netStartButton)

$bitLockerStatusButton = New-Object System.Windows.Forms.Button
$bitLockerStatusButton.Text = "Check BitLocker Status"
$bitLockerStatusButton.Location = New-Object System.Drawing.Point(10, 100)
$bitLockerStatusButton.Size = New-Object System.Drawing.Size(120, 23)
$bitLockerStatusButton.Add_Click({
    $outputTextBox.Clear()
    try {
        if (-not (Get-Command manage-bde -ErrorAction SilentlyContinue)) {
            $outputTextBox.Text = "Error: manage-bde is not available on this system."
            return
        }
        $bitLockerFeature = Get-WindowsOptionalFeature -Online -FeatureName BitLocker -ErrorAction SilentlyContinue
        if (-not $bitLockerFeature -or $bitLockerFeature.State -ne "Enabled") {
            $outputTextBox.Text = "Error: BitLocker feature is not enabled on this system."
            return
        }
        if (-not (Get-Volume -DriveLetter C -ErrorAction SilentlyContinue)) {
            $outputTextBox.Text = "Error: Drive C: is not a valid or accessible volume."
            return
        }
        $output = manage-bde -status C: 2>&1
        if ($LASTEXITCODE -eq 0) {
            $outputTextBox.Lines = $output
        } else {
            try {
                $bitLockerVolume = Get-BitLockerVolume -MountPoint C: -ErrorAction Stop
                $outputTextBox.Text = "BitLocker Status for C:`nProtection Status: $($bitLockerVolume.ProtectionStatus)`nEncryption Method: $($bitLockerVolume.EncryptionMethod)`nVolume Status: $($bitLockerVolume.VolumeStatus)"
            } catch {
                $outputTextBox.Text = "Error retrieving BitLocker status for C:. manage-bde error: $output`nGet-BitLockerVolume error: $($_.Exception.Message)"
            }
        }
    } catch {
        $outputTextBox.Text = "Error checking BitLocker status: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($bitLockerStatusButton, "Check BitLocker encryption status for drive C:")
$troubleGroupBox.Controls.Add($bitLockerStatusButton)

$bitLockerOffButton = New-Object System.Windows.Forms.Button
$bitLockerOffButton.Text = "Turn Off BitLocker"
$bitLockerOffButton.Location = New-Object System.Drawing.Point(140, 100)
$bitLockerOffButton.Size = New-Object System.Drawing.Size(120, 23)
$bitLockerOffButton.Add_Click({
    $outputTextBox.Clear()
    try {
        if (-not (Get-Command manage-bde -ErrorAction SilentlyContinue)) {
            $outputTextBox.Text = "Error: manage-bde is not available on this system."
            return
        }
        $bitLockerFeature = Get-WindowsOptionalFeature -Online -FeatureName BitLocker -ErrorAction SilentlyContinue
        if (-not $bitLockerFeature -or $bitLockerFeature.State -ne "Enabled") {
            $outputTextBox.Text = "Error: BitLocker feature is not enabled on this system."
            return
        }
        if (-not (Get-Volume -DriveLetter C -ErrorAction SilentlyContinue)) {
            $outputTextBox.Text = "Error: Drive C: is not a valid or accessible volume."
            return
        }
        $status = manage-bde -status C: 2>&1
        if ($status -match "Protection Status:\s+Protection On") {
            $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to turn off BitLocker on C:? This will decrypt the drive.", "Confirm Action", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($confirm -eq "Yes") {
                $output = manage-bde -off C: 2>&1
                if ($LASTEXITCODE -eq 0) {
                    $outputTextBox.Lines = $output
                    $outputTextBox.AppendText("`nBitLocker decryption initiated for C:. This may take some time.")
                } else {
                    $outputTextBox.Text = "Error turning off BitLocker on C:. Error: $output"
                }
            } else {
                $outputTextBox.Text = "BitLocker decryption cancelled."
            }
        } else {
            $outputTextBox.Text = "Error: BitLocker is not enabled on C: or drive is not encrypted."
        }
    } catch {
        $outputTextBox.Text = "Error turning off BitLocker: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($bitLockerOffButton, "Turn off BitLocker encryption for drive C: after confirmation")
$troubleGroupBox.Controls.Add($bitLockerOffButton)

$checkUpdatesButton = New-Object System.Windows.Forms.Button
$checkUpdatesButton.Text = "Check Windows Updates"
$checkUpdatesButton.Location = New-Object System.Drawing.Point(10, 175)
$checkUpdatesButton.Size = New-Object System.Drawing.Size(140, 23)
$checkUpdatesButton.Add_Click({
    $outputTextBox.Clear()
    try {
        $outputTextBox.Text = "Checking Windows Update history for the last 30 days...`n`n"
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        if ($historyCount -gt 0) {
            $updates = $searcher.QueryHistory(0, $historyCount) |
                Where-Object { $_.Date -gt (Get-Date).AddDays(-30) } |
                Select-Object Date, Title, @{Name="Result";Expression={switch($_.ResultCode){0{"Unknown"};1{"InProgress"};2{"Succeeded"};3{"SucceededWithErrors"};4{"Failed"};5{"Aborted"}}}}, @{Name="Operation";Expression={switch($_.Operation){1{"Installation"};2{"Uninstallation"};3{"Other"}}}} |
                Sort-Object Date -Descending
            if ($updates) {
                $index = 1
                foreach ($update in $updates) {
                    $outputTextBox.AppendText("Update $index`n")
                    $outputTextBox.AppendText("Date: $($update.Date)`n")
                    $outputTextBox.AppendText("Title: $($update.Title)`n")
                    $outputTextBox.AppendText("Result: $($update.Result)`n")
                    $outputTextBox.AppendText("Operation: $($update.Operation)`n`n")
                    $index++
                }
                $outputTextBox.AppendText("Windows Update history for the last 30 days listed above.")
            } else {
                $outputTextBox.Text += "No updates found in the last 30 days."
            }
        } else {
            $outputTextBox.Text += "No update history available."
        }
    } catch {
        $outputTextBox.Text = "Error checking updates: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($checkUpdatesButton, "List Windows Updates installed in the last 30 days")
$troubleGroupBox.Controls.Add($checkUpdatesButton)

$kbLabel = New-Object System.Windows.Forms.Label
$kbLabel.Text = "KB Number:"
$kbLabel.Location = New-Object System.Drawing.Point(10, 130)
$kbLabel.AutoSize = $true
$troubleGroupBox.Controls.Add($kbLabel)

$kbTextBox = New-Object System.Windows.Forms.TextBox
$kbTextBox.Location = New-Object System.Drawing.Point(10, 150)
$kbTextBox.Size = New-Object System.Drawing.Size(190, 20)
$kbTextBox.Add_KeyPress({
    if ($_.KeyChar -eq [char]13) {
        $uninstallKBButton.PerformClick()
    }
})
$troubleGroupBox.Controls.Add($kbTextBox)

$uninstallKBButton = New-Object System.Windows.Forms.Button
$uninstallKBButton.Text = "Uninstall KB"
$uninstallKBButton.Location = New-Object System.Drawing.Point(210, 150)
$uninstallKBButton.Size = New-Object System.Drawing.Size(90, 23)
$uninstallKBButton.Add_Click({
    $outputTextBox.Clear()
    $kbInput = $kbTextBox.Text.Trim()
    if (-not $kbInput) {
        $outputTextBox.Text = "Error: Please enter a KB number (e.g., KB1234567)."
        return
    }
    try {
        $kbNumber = if ($kbInput -match '^(KB)?(\d+)$') { $matches[2] } else { $null }
        if (-not $kbNumber) {
            $outputTextBox.Text = "Error: Invalid KB number format. Use KB1234567 or 1234567."
            return
        }
        $kbId = "KB$kbNumber"
        $hotfix = Get-HotFix -Id $kbId -ErrorAction SilentlyContinue
        if (-not $hotfix) {
            $outputTextBox.Text = "Error: $kbId is not installed on this system."
            return
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to uninstall $kbId? The system may require a restart.", "Confirm Uninstall", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -eq "Yes") {
            $outputTextBox.Text = "Uninstalling $kbId (this may take a moment)...`n"
            $process = Start-Process -FilePath "wusa.exe" -ArgumentList "/uninstall /kb:$kbNumber /quiet /norestart" -Wait -PassThru -NoNewWindow
            if ($process.ExitCode -eq 0) {
                $outputTextBox.Text = "$kbId uninstalled successfully. A restart may be required to complete the process."
            } elseif ($process.ExitCode -eq 3010) {
                $outputTextBox.Text = "$kbId uninstalled successfully. A restart is required to complete the process."
            } else {
                $outputTextBox.Text = "Error uninstalling $kbId. Exit code: $($process.ExitCode). Common codes: 2359302 (KB not installed), 3010 (restart required)."
            }
        } else {
            $outputTextBox.Text = "Uninstallation of $kbId cancelled."
        }
    } catch {
        $outputTextBox.Text = "Error uninstalling ${kbId}: $($_.Exception.Message)"
    }
})
$toolTip.SetToolTip($uninstallKBButton, "Uninstall the specified Windows Update KB number")
$troubleGroupBox.Controls.Add($uninstallKBButton)

# Create output text box
$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Location = New-Object System.Drawing.Point(10, 10)
$outputTextBox.Size = New-Object System.Drawing.Size(620, 690)
$outputTextBox.Multiline = $true
$outputTextBox.ScrollBars = "Vertical"
$outputTextBox.ReadOnly = $true
$outputTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputTextBox.WordWrap = $true  # Added to handle long text
$outputTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$outputPanel.Controls.Add($outputTextBox)

# Create a button to copy the output to clipboard
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = "Copy to Clipboard"
$copyButton.Location = New-Object System.Drawing.Point(10, 700)
$copyButton.Size = New-Object System.Drawing.Size(120, 23)
$copyButton.Add_Click({
    if ($outputTextBox.Text) {
        [System.Windows.Forms.Clipboard]::SetText($outputTextBox.Text)
        [System.Windows.Forms.MessageBox]::Show("Output copied to clipboard!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("No output to copy.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})
$outputPanel.Controls.Add($copyButton)

# Create a button to export the output to CSV
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export to CSV"
$exportButton.Location = New-Object System.Drawing.Point(140, 700)
$exportButton.Size = New-Object System.Drawing.Size(120, 23)
$exportButton.Add_Click({
    if ($outputTextBox.Text) {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $saveFileDialog.Title = "Save Output as CSV"
        if ($saveFileDialog.ShowDialog() -eq "OK") {
            try {
                $outputTextBox.Lines | ForEach-Object { 
                    [PSCustomObject]@{ Output = $_ } 
                } | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
                [System.Windows.Forms.MessageBox]::Show("Output exported to CSV!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error exporting to CSV: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No output to export.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})
$outputPanel.Controls.Add($exportButton)

# Resize event to reposition buttons
$form.Add_Resize({
    $copyButton.Location = New-Object System.Drawing.Point(10, ($outputTextBox.Bottom + 10))
    $exportButton.Location = New-Object System.Drawing.Point(140, ($outputTextBox.Bottom + 10))
})

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear All"
$clearButton.Location = New-Object System.Drawing.Point(10, 670)
$clearButton.Size = New-Object System.Drawing.Size(100, 23)
$clearButton.Add_Click({
    $userTextBox.Clear()
    $groupTextBox.Clear()
    $serviceTextBox.Clear()
    $appIdTextBox.Clear()
    $filePathTextBox.Clear()
    $kbTextBox.Clear()
    $outputTextBox.Clear()
})
$inputPanel.Controls.Add($clearButton)

# Device Name and Serial Number Labels
$deviceNameLabel = New-Object System.Windows.Forms.Label
$deviceNameLabel.Location = New-Object System.Drawing.Point(10, 700)
$deviceNameLabel.AutoSize = $true
try {
    $deviceName = Get-ComputerInfo | Select-Object -ExpandProperty CsName
    $deviceNameLabel.Text = "Device Name: $deviceName"
} catch {
    $deviceNameLabel.Text = "Device Name: Unknown"
}
$inputPanel.Controls.Add($deviceNameLabel)

$serialNumberLabel = New-Object System.Windows.Forms.Label
$serialNumberLabel.Location = New-Object System.Drawing.Point(10, 720)
$serialNumberLabel.AutoSize = $true
try {
    $serialNumber = Get-CimInstance -ClassName Win32_BIOS | Select-Object -ExpandProperty SerialNumber
    $serialNumberLabel.Text = "Serial Number: $serialNumber"
} catch {
    $serialNumberLabel.Text = "Serial Number: Unknown"
}
$inputPanel.Controls.Add($serialNumberLabel)

# Show the form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()