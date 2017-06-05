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
    $textHostName = "Host Name:`t`t`t$hostName`n"
    return $textHostName
}

function getIPAddr([string] $type) {
    $ipInfo = Get-NetIPAddress
    $textIPAddr = ""

    if ($type -eq "Ethernet") {
        $ipInfoEnthernet = $ipInfo | Where-Object {$_.InterfaceAlias -eq "Ethernet" -and $_.AddressFamily -eq "IPv4"}
        $ipAddrEthernet = $ipInfoEthernet.IPAddress
        $textIPAddr = "IP Address (Ethernet):`t`t$ipAddrEthernet"
    } elseif ($type -eq "Wi-Fi") {
        $ipInfoWiFi = $ipInfo | Where-Object {$_.InterfaceAlias -eq "Wi-Fi" -and $_.AddressFamily -eq "IPv4"}
        $ipAddrWiFi = $ipInfoWiFi.IPAddress
        $textIPAddr = "IP Address (Wi-Fi):`t`t$ipAddrWiFi"
    }

    return $textIPAddr
}

function getDriveInfo([string] $drive) {
    $driveInfo = Get-PSDrive | Where-Object {$_.Name -eq $drive}

    #convert from B to GB
    $driveUsed = $driveInfo.Used / 1GB
    $driveFree = $driveInfo.Free / 1GB

    $drivePercentFree = ([double] $driveFree) / (([double] $driveFree) + ([double] $driveUsed))
    $drivePercentFree = [math]::Round($drivePercentFree * 100, 2)
    
    $driveUsed = [math]::Round($driveUsed, 2)
    $driveFree = [math]::Round($driveFree, 2)

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
    $textUpTime = "Up Time:`t`t`t$upTime`n"

    $array = @($textCurrentTime, $textLastBootupTime, $textUpTime)
    return $array
}

function getProcessor() {
    $getProc = Get-WmiObject -Class Win32_Processor

    $systemName = $getProc.SystemName

    $numCores = $numCoresInfo.NumberOfCores

    $processorName = $getProc.Name

    $textSystemName = "System Name:`t`t`t$systemName"
    $textNumCores = "Number of Cores:`t`t$numCores"
    $textProcessorName = "Processor Name:`t`t`t$processorName"

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
    $aboutForm.Text = "About"
    $aboutForm.Font = $FONT
    $aboutForm.Size = New-Object Drawing.Size @(200, 125)
    $aboutForm.StartPosition = "CenterScreen"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false

    #title label
    $labelTitle = New-Object System.Windows.Forms.Label
    $labelTitle.Text = "Made by Kyle Spurlock"
    $labelTitle.Font = $FONT_BOLD
    $labelTitle.Size = New-Object System.Drawing.Size(200, 25)
    $labelTitle.Location = New-Object System.Drawing.Size(8, 25)
    $aboutForm.Controls.Add($labelTitle)

    #version label
    $labelVersion = New-Object System.Windows.Forms.Label
    $labelVersion.Text = "Version 1.0.0"
    $labelVersion.Size = New-Object System.Drawing.Size(200, 25)
    $labelVersion.Location = New-Object System.Drawing.Size(45, 50)
    $aboutForm.Controls.Add($labelVersion)

    $aboutForm.ShowDialog()
}

function save([System.Array] $labelTexts) {
    "Save" | Out-Host

    #save window
    $saveForm = New-Object Windows.Forms.Form
    $saveForm.Text = "Save"
    $saveForm.Font = $FONT
    $saveForm.Size = New-Object Drawing.Size @(275, 115)
    $saveForm.StartPosition = "CenterScreen"
    $saveForm.FormBorderStyle = "FixedDialog"
    $saveForm.MaximizeBox = $false

    #save label
    $labelFile = New-Object Windows.Forms.Label
    $labelFile.Text = "Enter Filename to Save to:"
    $labelFile.Font = $FONT_BOLD
    $labelFile.Size = New-Object Drawing.Size @(250, 50)
    $labelFile.Location = New-Object Drawing.Size @(5, 5)
    $saveForm.Controls.Add($labelFile)

    #save textbox
    $textboxFile = New-Object Windows.Forms.TextBox
    $textboxFile.Text = ""
    $textboxFile.Size = New-Object Drawing.Size @(200, 15)
    $textboxFile.Location = New-Object Drawing.Size @(5, 55)
    $saveForm.Controls.Add($textboxFile)

    #save button
    $buttonSave = New-Object Windows.Forms.Button
    $buttonSave.Text = "Save"
    $buttonSave.Add_Click({
        saveButtonClick $labelTexts $textboxFile.Text;
        $textboxFile.Text = "";
        $saveForm.Close()
    })
    $buttonSave.Size = New-Object Drawing.Size @(55, 25)
    $buttonSave.Location = New-Object Drawing.Size @(210, 55)
    $saveForm.Controls.Add($buttonSave)

    $saveForm.ShowDialog()
}

function saveButtonClick([System.Array] $labelTexts, [String] $fileLoc) {
    if ($fileLoc -eq "") {
        $fileLoc = "backup"
    } elseif ($fileLoc -match ".txt") {
        $fileLoc = $fileLoc.Substring(0, $fileLoc.Length - 4)
    }

    $fileLoc += ".txt"

    Set-Content "$PSScriptRoot\$fileLoc" $labelTexts
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
$form.Text = "Computer Information"

#set window font size
$form.Font = $FONT

#set size of window
$form.Size = New-Object Drawing.Size @(400, $WINDOW_SIZE_Y)

#set position of window
$form.StartPosition = "CenterScreen"

#prevent resizing of window
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$form.KeyPreview = $true
$form.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$form.Close()}})

$labels = @()

#menu
$menu = New-Object System.Windows.Forms.MenuStrip
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem

#region file
$menuFile.Text = "File"
$menu.Items.Add($menuFile)

$menuSave = New-Object System.Windows.Forms.ToolStripMenuItem
$menuSave.Text = "Save"
$menuSave.Add_Click({save $labelTexts})
$menuSave.ShortcutKeys = "Control, S"
$menuFile.DropDownItems.Add($menuSave)

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
$menuRefresh.ShortcutKeys = "Control, R"
$menuFile.DropDownItems.Add($menuRefresh)

$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExit.Text = "Exit"
$menuExit.Add_Click({$form.Close()})
$menuExit.ShortcutKeys = "Alt, F4"
$menuFile.DropDownItems.Add($menuExit)
#endregion

#region help
$menuHelp.Text = "Help"
$menu.Items.Add($menuHelp)

$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout.Text = "About"
$menuAbout.Add_Click({about})
$menuAbout.ShortcutKeys = "Control, A"
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