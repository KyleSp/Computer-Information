<#
    Made by Kyle Spurlock
#>

#-------------------------------------------------------
#constants
$NUM_LABEL_TEXTS = 12
$WINDOW_SIZE_Y = 400
$LABEL_SIZE_Y = 25
$FONT = New-Object System.Drawing.Font("Arial", 12)
$FONT_BOLD = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)

#-------------------------------------------------------
#functions

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

function getProcessor() {
    $getProc = Get-WmiObject -Class Win32_Processor

    $systemNameInfo = $getProc | Select SystemName | Out-String
    $systemName = $systemNameInfo.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)[2]
    $systemName = $systemName -replace "`n", ""

    $numCoresInfo = $getProc | Select NumberOfCores | Out-String
    $numCores = $numCoresInfo.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)[1]
    $numCores = $numCores -replace "`n", ""

    $processorNameInfo = $getProc | Select Name | Out-String
    $processorName = $processorNameInfo.Split("`n", [System.StringSplitOptions]::RemoveEmptyEntries)[3]
    $processorName = $processorName -replace "`n", ""

    $textSystemName = "System Name:`t`t`t$systemName`n"
    $textNumCores = "Number of Cores:`t`t$numCores`n"
    $textProcessorName = "Processor Name:`t`t`t$processorName`n"

    $array = @($textSystemName, $textNumCores, $textProcessorName)
    return $array
}

function getInfo([ref] $labelTexts, [boolean] $skipGetProcessor = $false) {
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

    #get processor info
    if (!$skipGetProcessor) {
        $processorInfo = getProcessor
        $labelTexts.Value[9] = $processorInfo[0]
        $labelTexts.Value[10] = $processorInfo[1]
        $labelTexts.Value[11] = $processorInfo[2]
    }
}

function refresh([ref] $labelTexts) {
    "Refresh" | Out-Host
    
    getInfo $labelTexts $true
}

function about() {
    "About" | Out-Host

    #about window
    $aboutForm = New-Object Windows.Forms.Form
    $aboutForm.text = "About"
    $aboutForm.Font = $FONT
    $aboutForm.Size = New-Object Drawing.Size @(200, 200)
    $aboutForm.StartPosition = "CenterScreen"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false

    #title label
    $labelTitle = New-Object System.Windows.Forms.Label
    $labelTitle.Text = "Made by Kyle Spurlock"
    $labelTitle.Font = $FONT_BOLD
    $labelTitle.Size = New-Object System.Drawing.Size(200, 25)
    $labelTitle.Location = New-Object System.Drawing.Size(10, 25)
    $aboutForm.Controls.Add($labelTitle)

    #version label
    $labelVersion = New-Object System.Windows.Forms.Label
    $labelVersion.Text = "Version 1.0.0"
    $labelVersion.Size = New-Object System.Drawing.Size(200, 25)
    $labelVersion.Location = New-Object System.Drawing.Size(45, 50)
    $aboutForm.Controls.Add($labelVersion)

    $aboutForm.ShowDialog()
}

#-------------------------------------------------------

$labelTexts = @("") * $NUM_LABEL_TEXTS

getInfo ([ref] $labelTexts)


#-------------------------------------------------------
#gui

Add-Type -AssemblyName System.Windows.Forms

#new window
$form = New-Object Windows.Forms.Form

#set window text
$form.text = "Computer Information"

#set window font size
$form.Font = $FONT

#set size of window
$form.Size = New-Object Drawing.Size @(400, $WINDOW_SIZE_Y)

#set position of window
$form.StartPosition = "CenterScreen"

#prevent resizing of window
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

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
    #$loc = $i * $LABEL_SIZE_Y + 35
    $loc = 35
    for ($j = 0; $j -lt $i * 2; ++$j) {
        if ($j % 2 -ne 0) {
            $loc += $labels[$j].Size.Height
        }
    }

    $sizeMult = [int] ($label2.Text.Length / 20)
    if ($sizeMult -eq 0) {
        $sizeMult = 1
    }

    $size = $LABEL_SIZE_Y * $sizeMult

    $label1.Size = New-Object System.Drawing.Size(200, $size)
    $label1.Location = New-Object System.Drawing.Size(10, $loc)
    $label2.Size = New-Object System.Drawing.Size(200, $size)
    $label2.Location = New-Object System.Drawing.Size(210, $loc)

    #add labels to window
    $labels += $label1
    $labels += $label2

    $form.Controls.Add($label1)
    $form.Controls.Add($label2)
}

#display window
$form.ShowDialog()