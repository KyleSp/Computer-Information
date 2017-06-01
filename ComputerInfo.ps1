<#
    Made by Kyle Spurlock
#>

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
    $drivePercentFree = [math]::Round($drivePercentFree, 2)
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

function refresh() {
    "Refresh" | Out-Host
}

#-------------------------------------------------------

$text = ""

#host name
$hostName = getHostName

#ip address (ethernet)
$textIPAddrEthernet = getIPAddr "Ethernet"

#ip address (wi-fi)
$textIPAddrWiFi = getIPAddr "Wi-Fi"

#C drive info
$driveInfo = getDriveInfo "C"
$textDriveUsed = $driveInfo[0]
$textDriveFree = $driveInfo[1]
$textDrivePercentFree = $driveInfo[2]

#system uptime
$upTimeInfo = getSystemUpTime
$textCurrentTime = $upTimeInfo[0]
$textLastBootupTime = $upTimeInfo[1]
$textUpTime = $upTimeInfo[2]

#$text | Out-Host
#$text | Add-Content $PSScriptRoot\ComputerInfo.txt


#-------------------------------------------------------
#make gui

Add-Type -AssemblyName System.Windows.Forms

#new window
$form = New-Object Windows.Forms.Form

#set window text
$form.text = "Computer Information"

#set window font size
$form.Font = New-Object System.Drawing.Font("Arial", 12)

#set size of window
$form.Size = New-Object Drawing.Size @(400, 300)

#set position of window
$form.StartPosition = "CenterScreen"

#prevent resizing of window
$form.FormBorderStyle = "FixedDialog"

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
$menuRefresh.Add_Click({refresh})
$menuFile.DropDownItems.Add($menuRefresh)
#endregion

#region help
$menuHelp.Text = "Help"
$menu.Items.Add($menuHelp)

$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout.Text = "About"
$menuAbout.Add_Click({"About" | Out-Host})
$menuHelp.DropDownItems.Add($menuAbout)
#endregion

#add menu to window
$form.Controls.Add($menu)

Clear-Host

#add labels
$labelTexts = @(
    $textHostName,
    $textIPAddrEthernet,
    $textIPAddrWiFi,
    $textDriveUsed,
    $textDriveFree,
    $textDrivePercentFree,
    $textCurrentTime,
    $textLastBootupTime,
    $textUpTime
)

for ($i = 0; $i -lt $labelTexts.Length; ++$i) {
    #make labels
    $label1 = New-Object System.Windows.Forms.Label
    $label2 = New-Object System.Windows.Forms.Label

    #format text
    $index = $labelTexts[$i].IndexOf("`t")
    $text = $labelTexts[$i] -replace "`t", ""
    $label1.Text = $text.Substring(0, $index)
    $label2.Text = $text.Substring($index)

    #set label location and size
    $loc = $i * 25 + 35
    $label1.Size = New-Object System.Drawing.Size(200, 25)
    $label1.Location = New-Object System.Drawing.Size(10, $loc)
    $label2.Size = New-Object System.Drawing.Size(200, 25)
    $label2.Location = New-Object System.Drawing.Size(210, $loc)

    #add labels to window
    $form.Controls.Add($label1)
    $form.Controls.Add($label2)
}

#display window
$form.ShowDialog()