$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# Define variables
$ApplicationName = 'PingCastle'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$PingCastlePath = Join-Path $ScriptRoot $ApplicationName
$ExecutablePath = Join-Path $PingCastlePath "$ApplicationName.exe"
$ReportFolder = Join-Path $PingCastlePath "Reports"
$ReportFileNameFormat = "ad_hc_{0}_{1}.html" # Domain and Date will be added dynamically

# Function to extract score from XML report
function Get-PingCastleScores {
    param (
        [string]$xmlReportPath
    )

    $xmlContent = [xml](Get-Content $xmlReportPath)
    $scores = @{
        "GlobalScore" = $xmlContent.SelectSingleNode("/HealthcheckData/GlobalScore").InnerText
        "StaleObjectsScore" = $xmlContent.SelectSingleNode("/HealthcheckData/StaleObjectsScore").InnerText
        "PrivilegedGroupScore" = $xmlContent.SelectSingleNode("/HealthcheckData/PrivilegiedGroupScore").InnerText
        "TrustScore" = $xmlContent.SelectSingleNode("/HealthcheckData/TrustScore").InnerText
        "AnomalyScore" = $xmlContent.SelectSingleNode("/HealthcheckData/AnomalyScore").InnerText
    }
    return $scores
}

# Function to compare two sets of scores
function Compare-Scores {
    param (
        [hashtable]$CurrentScores,
        [hashtable]$PreviousScores
    )

    foreach ($key in $CurrentScores.Keys) {
        if ($PreviousScores[$key] -ne $CurrentScores[$key]) {
            return $true
        }
    }
    return $false
}

# Function to run PingCastle and generate report
function Run-PingCastle {
    if (-not (Test-Path $ExecutablePath)) {
        throw "Executable not found: $ExecutablePath"
    }

    if (-not (Test-Path $ReportFolder)) {
        New-Item -Path $ReportFolder -ItemType Directory | Out-Null
    }

    try {
        Set-Location -Path $PingCastlePath
        Start-Process -FilePath $ExecutablePath -ArgumentList "--healthcheck --level Full" -WindowStyle Hidden -Wait
    } catch {
        throw "Failed to execute PingCastle: $_"
    }

    # Adjust the report file name and location
    $DefaultReportName = "ad_hc_{0}.html" -f $env:USERDNSDOMAIN.ToLower()
    $DefaultReportPath = Join-Path $PingCastlePath $DefaultReportName
    $NewReportFileName = $ReportFileNameFormat -f $env:USERDNSDOMAIN.ToLower(), (Get-Date -UFormat "%d%m%y_%H%M%S")
    $NewReportPath = Join-Path $ReportFolder $NewReportFileName

    if (Test-Path $DefaultReportPath) {
        Move-Item -Path $DefaultReportPath -Destination $NewReportPath
    } else {
        throw "Report not generated: $DefaultReportPath"
    }

    return $NewReportPath
}

# Function to create a zip file of the report
function Compress-ZipFile {
    param (
        [string]$FilePath
    )

    $zipFilePath = "$FilePath.zip"
    if (Test-Path $zipFilePath) { Remove-Item $zipFilePath } # Remove existing zip file if any

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

    $zip = [System.IO.Compression.ZipFile]::Open($zipFilePath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $fileName = [System.IO.Path]::GetFileName($FilePath)
        $fileInfo = New-Object System.IO.FileInfo($FilePath)
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $fileInfo.FullName, $fileName, [System.IO.Compression.CompressionLevel]::Optimal)
    } finally {
        $zip.Dispose()
    }

    return $zipFilePath
}

# Function to send email with attachment
function Send-EmailReport {
    param (
        [string]$ZipFilePath,
        [string]$RecipientEmail,
        [bool]$IsDataChanged
    )

    $currentMonthYear = Get-Date -UFormat "%b %Y"

    # Email sending parameters
    $smtpServer = "" # Insert SMTP Server
    $smtpPort = 587
    $smtpUser = "" # Insert SMTP User
    $smtpPassword = "" # Insert SMTP Password

    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.From = "" # Insert SMTP Sender Email
    $mailMessage.To.Add($RecipientEmail)
    if ($IsDataChanged) {
        $mailMessage.Subject = "$currentMonthYear Monthly PingCastle Report (Changes present!)"
        $mailMessage.Body = "Please find the attached updated PingCastle report. There are some changes this month."
    } else {
        $mailMessage.Subject = "$currentMonthYear Monthly PingCastle Report (No changes)"
        $mailMessage.Body = "No changes detected in the latest PingCastle report."
    }

    try {
        # Create and add the attachment
        $attachment = New-Object System.Net.Mail.Attachment($ZipFilePath)
        $mailMessage.Attachments.Add($attachment)

        $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
        $smtpClient.EnableSsl = $true # Enable SSL/TLS
        $smtpClient.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPassword)

        $smtpClient.Send($mailMessage)
        Write-Host "Email sent with report to $RecipientEmail"
    } catch {
        Write-Error "Failed to send email: $_"
    } finally {
        $mailMessage.Dispose()
        if (Test-Path $ZipFilePath) { Remove-Item $ZipFilePath } # Delete the zip file after sending
    }
}

# Main script execution
try {
    $NewReportPath = Run-PingCastle
    Create-ZipFile -FilePath $NewReportPath
    $ZipFilePath = "$NewReportPath.zip"

    # Dynamically determine the XML file name based on the domain name
    $domainName = $env:USERDNSDOMAIN.ToLower()
    $xmlFilePath = Join-Path $PingCastlePath "ad_hc_$domainName.xml"
    $oldXmlFilePath = Join-Path $PingCastlePath "ad_hc_${domainName}_old.xml"

    # Compare Scores
    $IsDataChanged = $false
    if (Test-Path $oldXmlFilePath) {
        $currentScores = Get-PingCastleScores -xmlReportPath $xmlFilePath
        $previousScores = Get-PingCastleScores -xmlReportPath $oldXmlFilePath
        $IsDataChanged = Compare-Scores -CurrentScores $currentScores -PreviousScores $previousScores
    }

    # Send email report (Add in recepient email)
    Send-EmailReport -ZipFilePath $ZipFilePath -RecipientEmail "" -IsDataChanged $IsDataChanged

    # Rename the new XML file and delete the old one
    if (Test-Path $xmlFilePath) {
        if (Test-Path $oldXmlFilePath) {
            Remove-Item -Path $oldXmlFilePath
        }
        $newName = "ad_hc_${domainName}_old.xml"
        Rename-Item -Path $xmlFilePath -NewName $newName
    }
} catch {
    Write-Error $_
}

# Update Logic
$UpdaterName = "PingCastleAutoUpdater"
$UpdaterPath = Join-Path $PingCastlePath "$UpdaterName.exe"

if (Test-Path $UpdaterPath) {
    try {
        Set-Location -Path $PingCastlePath
        Start-Process -FilePath $UpdaterPath -ArgumentList "--wait-for-days 30" -WindowStyle Hidden -Wait
        Write-Host "PingCastle updated successfully."
    } catch {
        Write-Error "Failed to run the updater: $_"
    }
} else {
    Write-Host "Updater not found: $UpdaterPath"
}