<#
    Made by Kyle Spurlock
#>

$FONT = New-Object System.Drawing.Font("Arial", 12)
$FONT_BOLD = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)

#-------------------------------------------------------

function getHostName() {
    $hostName = $env:USERNAME
    $textHostName = "Host Name:`t`t`t`t$hostName`n"
    return $textHostName
}

function getIPAddr([string] $type) {
    $ipInfo = Get-NetIPAddress
    
    if ($type -eq "Ethernet") {
        $ipInfoEthernet = $ipInfo | Where-Object {$_.InterfaceAlias -eq "Ethernet" -and $_.AddressFamily -eq "IPv4"} | Select IPAddress | Out-String
        $ipAddrEthernet = $ipInfoEthernet.Split((" ", "`n"), [System.StringSplitOptions]::RemoveEmptyEntries)[5]
        $textIPAddrEthernet = "IP Address (Ethernet):`t$ipAddrEthernet`n"
        return $textIPAddrEthernet
    } elseif ($type -eq "Wi-Fi") {
        $ipInfoWiFi = $ipInfo | Where-Object {$_.InterfaceAlias -eq "Wi-Fi" -and $_.AddressFamily -eq "IPv4"} | Select IPAddress | Out-String
        $ipAddrWiFi = $ipInfoWiFi.Split((" ", "`n"), [System.StringSplitOptions]::RemoveEmptyEntries)[5]
        $textIPAddrWiFi = "IP Address (Wi-Fi):`t`t$ipAddrWiFi`n"
        return $textIPAddrWiFi
    } else {
        return ""
    }
}

function getDriveInfo([string] $drive) {
    $driveInfo = Get-PSDrive | Where-Object {$_.Name -eq $drive} | Out-String
    $driveUsed = $driveInfo.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)[13]
    $driveFree = $driveInfo.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)[14]
    $drivePercentFree = ([double] $driveFree) / (([double] $driveFree) + ([double] $driveUsed))
    $drivePercentFree = [math]::Round($drivePercentFree * 100, 2)
    $textDriveUsed = "$drive Drive (Used):`t`t`t$driveUsed GB`n"
    $textDriveFree = "$drive Drive (Free):`t`t`t$driveFree GB`n"
    $textDrivePercentFree = "$drive Drive % Free:`t`t`t$drivePercentFree %`n"

    $array = @($textDriveUsed, $textDriveFree, $textDrivePercentFree)
    return $array
}

function getSystemUpTime() {
    $opSysInst = Get-CimInstance Win32_OperatingSystem
    $currentTime = $opSysInst.LocalDateTime
    $lastBootupTime = $opSysInst.LastBootUpTime
    $upTime = ($currentTime - $lastBootUpTime).ToString()
    $upTime = $uptime.Substring(0, $uptime.length - 8)
    $textCurrentTime = "Current Time:`t`t`t$currentTime`n"
    $textLastBootupTime = "Last Bootup Time:`t`t$lastBootupTime`n"
    $textUpTime = "Up Time:`t`t`t`t$upTime`n"

    $array = @($textCurrentTime, $textLastBootupTime, $textUpTime)
    return $array
}

function getInfo([ref] $labelTexts) {
    #host name
    $labelTexts.Value[0] = getHostName

    #ip address (ethernet)
    $labelTexts.Value[1] = getIPAddr "Ethernet"

    #ip address (wi-fi)
    $labelTexts.Value[2] = getIPAddr "Wi-Fi"

    #C drive info
    $driveInfo = getDriveInfo "C"
    $labelTexts.Value[3] = $driveInfo[0]
    $labelTexts.Value[4] = $driveInfo[1]
    $labelTexts.Value[5] = $driveInfo[2]

    #system uptime
    $upTimeInfo = getSystemUpTime
    $labelTexts.Value[6] = $upTimeInfo[0]
    $labelTexts.Value[7] = $upTimeInfo[1]
    $labelTexts.Value[8] = $upTimeInfo[2]
}

function refresh([ref] $labelTexts) {
    "Refresh" | Out-Host
    
    getInfo ($labelTexts)
}

function about() {
    "About" | Out-Host

    #make about window
    $aboutForm = New-Object Windows.Forms.Form
    $aboutForm.text = "About"
    $aboutForm.Font = $FONT
    $aboutForm.Size = New-Object Drawing.Size @(200, 200)
    $aboutForm.StartPosition = "CenterScreen"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false

    #make title label
    $labelTitle = New-Object System.Windows.Forms.Label
    $labelTitle.Text = "Made by Kyle Spurlock"
    $labelTitle.Font = $FONT_BOLD
    $labelTitle.Size = New-Object System.Drawing.Size(200, 25)
    $labelTitle.Location = New-Object System.Drawing.Size(10, 25)
    $aboutForm.Controls.Add($labelTitle)

    #make version label
    $labelVersion = New-Object System.Windows.Forms.Label
    $labelVersion.Text = "Version 1.0.0"
    $labelVersion.Size = New-Object System.Drawing.Size(200, 25)
    $labelVersion.Location = New-Object System.Drawing.Size(45, 50)
    $aboutForm.Controls.Add($labelVersion)

    $aboutForm.ShowDialog()
}

#-------------------------------------------------------

$labelTexts = @("") * 9

getInfo ([ref] $labelTexts)


#-------------------------------------------------------
#make gui

Add-Type -AssemblyName System.Windows.Forms

#new window
$form = New-Object Windows.Forms.Form

#set window text
$form.text = "Computer Information"

#set window font size
$form.Font = $FONT

#set size of window
$form.Size = New-Object Drawing.Size @(400, 300)

#set position of window
$form.StartPosition = "CenterScreen"

#prevent resizing of window
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

#$labels = @() * 18
$labels = @()

#menu
$menu = New-Object System.Windows.Forms.MenuStrip
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem

#region file
$menuFile.Text = "File"
$menu.Items.Add($menuFile)

$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExit.Text = "Exit"
$menuExit.Add_Click({$form.Close()})
$menuFile.DropDownItems.Add($menuExit)

$menuRefresh = New-Object System.Windows.Forms.ToolStripMenuItem
$menuRefresh.Text = "Refresh"
$menuRefresh.Add_Click({
    refresh ([ref] $labelTexts)

    for ($i = 0; $i -lt $labelTexts.Length; ++$i) {
        $index = $labelTexts[$i].IndexOf("`t")
        $text = $labelTexts[$i] -replace "`t", ""
        $labels[$i * 2 + 1].Text = $text.Substring($index)
    }
})
$menuFile.DropDownItems.Add($menuRefresh)
#endregion

#region help
$menuHelp.Text = "Help"
$menu.Items.Add($menuHelp)

$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout.Text = "About"
$menuAbout.Add_Click({about})
$menuHelp.DropDownItems.Add($menuAbout)
#endregion

#add menu to window
$form.Controls.Add($menu)

Clear-Host

#add labels
for ($i = 0; $i -lt $labelTexts.Length; ++$i) {
    #make labels
    $label1 = New-Object System.Windows.Forms.Label
    $label2 = New-Object System.Windows.Forms.Label

    #format text
    $index = $labelTexts[$i].IndexOf("`t")
    $text = $labelTexts[$i] -replace "`t", ""
    $label1.Text = $text.Substring(0, $index)
    $label2.Text = $text.Substring($index)

    $label1.Font = $FONT_BOLD

    #set label location and size
    $loc = $i * 25 + 35
    $label1.Size = New-Object System.Drawing.Size(200, 25)
    $label1.Location = New-Object System.Drawing.Size(10, $loc)
    $label2.Size = New-Object System.Drawing.Size(200, 25)
    $label2.Location = New-Object System.Drawing.Size(210, $loc)

    #add labels to window
    $labels += $label1
    $labels += $label2

    $form.Controls.Add($label1)
    $form.Controls.Add($label2)
}

#display window
$form.ShowDialog()