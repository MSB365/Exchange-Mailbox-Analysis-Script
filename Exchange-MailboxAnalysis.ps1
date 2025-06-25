<#
.SYNOPSIS
    Exchange Mailbox Analysis Script with EWS Integration
.DESCRIPTION
    This script analyzes Exchange mailboxes using EWS API and generates comprehensive reports
.AUTHOR
    Exchange Admin Tool
.VERSION
    1.0
#>

# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import Exchange Management Shell if not already loaded
if (!(Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
    Write-Host "Loading Exchange Management Shell..." -ForegroundColor Yellow
    try {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        Write-Host "Exchange Management Shell loaded successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to load Exchange Management Shell. Please run this script from Exchange Management Shell."
        exit 1
    }
}

# Function to show file dialog for CSV selection
function Select-CSVFile {
    Write-Host "Opening file dialog for CSV selection..." -ForegroundColor Cyan
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $OpenFileDialog.Title = "Select CSV file with mailbox list"
    
    if ($OpenFileDialog.ShowDialog() -eq 'OK') {
        Write-Host "Selected file: $($OpenFileDialog.FileName)" -ForegroundColor Green
        return $OpenFileDialog.FileName
    }
    else {
        Write-Host "No file selected. Exiting..." -ForegroundColor Red
        exit 1
    }
}

# Function to validate CSV structure
function Test-CSVStructure {
    param([string]$FilePath)
    
    Write-Host "Validating CSV structure..." -ForegroundColor Cyan
    
    try {
        $csvData = Import-Csv $FilePath
        if ($csvData.Count -eq 0) {
            throw "CSV file is empty"
        }
        
        # Check for required columns (flexible - can be EmailAddress, UserPrincipalName, or SamAccountName)
        $headers = $csvData[0].PSObject.Properties.Name
        $validHeaders = @('EmailAddress', 'UserPrincipalName', 'SamAccountName', 'Identity', 'Mailbox')
        
        $foundHeader = $headers | Where-Object { $_ -in $validHeaders }
        if (-not $foundHeader) {
            throw "CSV must contain at least one of these columns: $($validHeaders -join ', ')"
        }
        
        Write-Host "CSV validation successful. Found $($csvData.Count) entries." -ForegroundColor Green
        return $csvData, $foundHeader[0]
    }
    catch {
        Write-Error "CSV validation failed: $($_.Exception.Message)"
        exit 1
    }
}

# Function to create EWS service
function New-EWSService {
    param(
        [string]$EmailAddress,
        [string]$ExchangeVersion = "Exchange2013_SP1"
    )
    
    try {
        # Load EWS Managed API
        $EWSPath = "${env:ProgramFiles}\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
        if (-not (Test-Path $EWSPath)) {
            $EWSPath = "${env:ProgramFiles(x86)}\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
        }
        
        if (Test-Path $EWSPath) {
            Add-Type -Path $EWSPath
        }
        else {
            throw "EWS Managed API not found. Please install Exchange Web Services Managed API 2.2"
        }
        
        # Create EWS service
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion)
        $service.UseDefaultCredentials = $true
        $service.AutodiscoverUrl($EmailAddress)
        
        return $service
    }
    catch {
        Write-Warning "Failed to create EWS service for $EmailAddress : $($_.Exception.Message)"
        return $null
    }
}

# Function to get mailbox statistics via EWS
function Get-MailboxEWSStats {
    param(
        [string]$EmailAddress,
        [object]$EWSService
    )
    
    $stats = @{
        EmailAddress = $EmailAddress
        TotalMessages = 0
        ReadMessages = 0
        UnreadMessages = 0
        LastReceivedDate = $null
        LastSentDate = $null
        LastReadMessages = @()
        Error = $null
    }
    
    try {
        if (-not $EWSService) {
            $stats.Error = "EWS Service not available"
            return $stats
        }
        
        # Get Inbox folder
        $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWSService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
        $stats.TotalMessages = $inbox.TotalCount
        $stats.UnreadMessages = $inbox.UnreadCount
        $stats.ReadMessages = $inbox.TotalCount - $inbox.UnreadCount
        
        # Get recent messages for last received date
        $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(50)
        $itemView.OrderBy.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [Microsoft.Exchange.WebServices.Data.SortDirection]::Descending)
        
        $findResults = $EWSService.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $itemView)
        if ($findResults.Items.Count -gt 0) {
            $stats.LastReceivedDate = $findResults.Items[0].DateTimeReceived
        }
        
        # Get last read messages (last 5)
        $readMessages = $findResults.Items | Where-Object { $_.IsRead -eq $true } | Select-Object -First 5
        foreach ($msg in $readMessages) {
            $stats.LastReadMessages += @{
                Subject = $msg.Subject
                DateTimeReceived = $msg.DateTimeReceived
                From = if ($msg.From) { $msg.From.Address } else { "Unknown" }
            }
        }
        
        # Get Sent Items folder for last sent date
        try {
            $sentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWSService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems)
            $sentItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)
            $sentItemView.OrderBy.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeSent, [Microsoft.Exchange.WebServices.Data.SortDirection]::Descending)
            
            $sentResults = $EWSService.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems, $sentItemView)
            if ($sentResults.Items.Count -gt 0) {
                $stats.LastSentDate = $sentResults.Items[0].DateTimeSent
            }
        }
        catch {
            Write-Warning "Could not access Sent Items for $EmailAddress"
        }
    }
    catch {
        $stats.Error = $_.Exception.Message
        Write-Warning "Error getting EWS stats for $EmailAddress : $($_.Exception.Message)"
    }
    
    return $stats
}

# Function to get mailbox permissions
function Get-MailboxPermissions {
    param([string]$Identity)
    
    $permissions = @{
        FullAccess = @()
        SendAs = @()
        SendOnBehalf = @()
    }
    
    try {
        # Get FullAccess permissions
        $fullAccessPerms = Get-MailboxPermission -Identity $Identity | Where-Object { $_.AccessRights -contains "FullAccess" -and $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY\SELF" }
        foreach ($perm in $fullAccessPerms) {
            $permissions.FullAccess += "$($perm.User) ($($perm.AccessRights -join ', '))"
        }
        
        # Get SendAs permissions
        $sendAsPerms = Get-ADPermission -Identity $Identity | Where-Object { $_.ExtendedRights -contains "Send-As" -and $_.IsInherited -eq $false }
        foreach ($perm in $sendAsPerms) {
            $permissions.SendAs += "$($perm.User)"
        }
        
        # Get SendOnBehalf permissions
        $mailbox = Get-Mailbox -Identity $Identity
        if ($mailbox.GrantSendOnBehalfTo) {
            foreach ($user in $mailbox.GrantSendOnBehalfTo) {
                $permissions.SendOnBehalf += $user.ToString()
            }
        }
    }
    catch {
        Write-Warning "Error getting permissions for $Identity : $($_.Exception.Message)"
    }
    
    return $permissions
}

# Function to generate HTML report
function New-HTMLReport {
    param(
        [array]$Results,
        [string]$OutputPath
    )
    
    Write-Host "Generating HTML report..." -ForegroundColor Cyan
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Mailbox Analysis Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .header { background-color: #0078d4; color: white; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
        .summary { background-color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .mailbox-card { background-color: white; margin-bottom: 20px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
        .mailbox-header { background-color: #106ebe; color: white; padding: 15px; font-weight: bold; }
        .mailbox-content { padding: 15px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 15px; }
        .stat-item { background-color: #f8f9fa; padding: 10px; border-radius: 3px; border-left: 4px solid #0078d4; }
        .stat-label { font-weight: bold; color: #666; font-size: 12px; text-transform: uppercase; }
        .stat-value { font-size: 18px; color: #333; margin-top: 5px; }
        .permissions { margin-top: 15px; }
        .permission-type { background-color: #e3f2fd; padding: 10px; margin: 5px 0; border-radius: 3px; }
        .recent-messages { margin-top: 15px; }
        .message-item { background-color: #f8f9fa; padding: 8px; margin: 3px 0; border-radius: 3px; font-size: 12px; }
        .error { background-color: #ffebee; color: #c62828; padding: 10px; border-radius: 3px; }
        .timestamp { color: #666; font-size: 11px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Exchange Mailbox Analysis Report</h1>
        <p>Generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    </div>
    
    <div class="summary">
        <h2>Summary</h2>
        <p><strong>Total Mailboxes Analyzed:</strong> $($Results.Count)</p>
        <p><strong>Successful Analyses:</strong> $(($Results | Where-Object { -not $_.Error }).Count)</p>
        <p><strong>Failed Analyses:</strong> $(($Results | Where-Object { $_.Error }).Count)</p>
    </div>
"@

    foreach ($result in $Results) {
        $html += @"
    <div class="mailbox-card">
        <div class="mailbox-header">$($result.EmailAddress)</div>
        <div class="mailbox-content">
"@
        
        if ($result.Error) {
            $html += @"
            <div class="error">
                <strong>Error:</strong> $($result.Error)
            </div>
"@
        }
        else {
            $html += @"
            <div class="stats-grid">
                <div class="stat-item">
                    <div class="stat-label">Total Messages</div>
                    <div class="stat-value">$($result.TotalMessages)</div>
                </div>
                <div class="stat-item">
                    <div class="stat-label">Read Messages</div>
                    <div class="stat-value">$($result.ReadMessages)</div>
                </div>
                <div class="stat-item">
                    <div class="stat-label">Unread Messages</div>
                    <div class="stat-value">$($result.UnreadMessages)</div>
                </div>
                <div class="stat-item">
                    <div class="stat-label">Last Received</div>
                    <div class="stat-value">$(if ($result.LastReceivedDate) { $result.LastReceivedDate.ToString("yyyy-MM-dd HH:mm") } else { "N/A" })</div>
                </div>
                <div class="stat-item">
                    <div class="stat-label">Last Sent</div>
                    <div class="stat-value">$(if ($result.LastSentDate) { $result.LastSentDate.ToString("yyyy-MM-dd HH:mm") } else { "N/A" })</div>
                </div>
            </div>
            
            <div class="permissions">
                <h4>Mailbox Permissions</h4>
"@
            
            if ($result.Permissions.FullAccess.Count -gt 0) {
                $html += @"
                <div class="permission-type">
                    <strong>Full Access:</strong><br>
                    $($result.Permissions.FullAccess -join '<br>')
                </div>
"@
            }
            
            if ($result.Permissions.SendAs.Count -gt 0) {
                $html += @"
                <div class="permission-type">
                    <strong>Send As:</strong><br>
                    $($result.Permissions.SendAs -join '<br>')
                </div>
"@
            }
            
            if ($result.Permissions.SendOnBehalf.Count -gt 0) {
                $html += @"
                <div class="permission-type">
                    <strong>Send On Behalf:</strong><br>
                    $($result.Permissions.SendOnBehalf -join '<br>')
                </div>
"@
            }
            
            if ($result.Permissions.FullAccess.Count -eq 0 -and $result.Permissions.SendAs.Count -eq 0 -and $result.Permissions.SendOnBehalf.Count -eq 0) {
                $html += "<p><em>No additional permissions found</em></p>"
            }
            
            $html += "</div>"
            
            if ($result.LastReadMessages.Count -gt 0) {
                $html += @"
            <div class="recent-messages">
                <h4>Last 5 Read Messages</h4>
"@
                foreach ($msg in $result.LastReadMessages) {
                    $html += @"
                <div class="message-item">
                    <strong>$($msg.Subject)</strong><br>
                    <span class="timestamp">From: $($msg.From) | Received: $($msg.DateTimeReceived.ToString("yyyy-MM-dd HH:mm"))</span>
                </div>
"@
                }
                $html += "</div>"
            }
        }
        
        $html += @"
        </div>
    </div>
"@
    }
    
    $html += @"
</body>
</html>
"@
    
    try {
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "HTML report saved to: $OutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to save HTML report: $($_.Exception.Message)"
    }
}

# Function to generate CSV report
function New-CSVReport {
    param(
        [array]$Results,
        [string]$OutputPath
    )
    
    Write-Host "Generating CSV report..." -ForegroundColor Cyan
    
    $csvData = @()
    
    foreach ($result in $Results) {
        $csvRow = [PSCustomObject]@{
            EmailAddress = $result.EmailAddress
            TotalMessages = $result.TotalMessages
            ReadMessages = $result.ReadMessages
            UnreadMessages = $result.UnreadMessages
            LastReceivedDate = if ($result.LastReceivedDate) { $result.LastReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
            LastSentDate = if ($result.LastSentDate) { $result.LastSentDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
            FullAccessPermissions = ($result.Permissions.FullAccess -join "; ")
            SendAsPermissions = ($result.Permissions.SendAs -join "; ")
            SendOnBehalfPermissions = ($result.Permissions.SendOnBehalf -join "; ")
            LastReadMessagesCount = $result.LastReadMessages.Count
            Error = $result.Error
        }
        $csvData += $csvRow
    }
    
    try {
        $csvData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV report saved to: $OutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to save CSV report: $($_.Exception.Message)"
    }
}

# Main execution
Write-Host "=== Exchange Mailbox Analysis Script ===" -ForegroundColor Magenta
Write-Host "Starting mailbox analysis process..." -ForegroundColor Yellow

# Step 1: Select CSV file
$csvFile = Select-CSVFile

# Step 2: Validate CSV structure
$csvData, $identityColumn = Test-CSVStructure -FilePath $csvFile

# Step 3: Process each mailbox
Write-Host "`nProcessing $($csvData.Count) mailboxes..." -ForegroundColor Yellow
$results = @()
$counter = 0

foreach ($row in $csvData) {
    $counter++
    $identity = $row.$identityColumn
    
    Write-Host "`n[$counter/$($csvData.Count)] Processing: $identity" -ForegroundColor Cyan
    
    try {
        # Get mailbox information
        Write-Host "  - Getting mailbox information..." -ForegroundColor Gray
        $mailbox = Get-Mailbox -Identity $identity -ErrorAction Stop
        $emailAddress = $mailbox.PrimarySmtpAddress.ToString()
        
        # Create EWS service
        Write-Host "  - Connecting to EWS..." -ForegroundColor Gray
        $ewsService = New-EWSService -EmailAddress $emailAddress
        
        # Get EWS statistics
        Write-Host "  - Retrieving mailbox statistics..." -ForegroundColor Gray
        $ewsStats = Get-MailboxEWSStats -EmailAddress $emailAddress -EWSService $ewsService
        
        # Get permissions
        Write-Host "  - Checking permissions..." -ForegroundColor Gray
        $permissions = Get-MailboxPermissions -Identity $identity
        
        # Compile results
        $result = [PSCustomObject]@{
            EmailAddress = $emailAddress
            TotalMessages = $ewsStats.TotalMessages
            ReadMessages = $ewsStats.ReadMessages
            UnreadMessages = $ewsStats.UnreadMessages
            LastReceivedDate = $ewsStats.LastReceivedDate
            LastSentDate = $ewsStats.LastSentDate
            LastReadMessages = $ewsStats.LastReadMessages
            Permissions = $permissions
            Error = $ewsStats.Error
        }
        
        $results += $result
        
        if ($ewsStats.Error) {
            Write-Host "  - Completed with errors" -ForegroundColor Yellow
        }
        else {
            Write-Host "  - Completed successfully" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  - Failed: $($_.Exception.Message)" -ForegroundColor Red
        
        $result = [PSCustomObject]@{
            EmailAddress = $identity
            TotalMessages = 0
            ReadMessages = 0
            UnreadMessages = 0
            LastReceivedDate = $null
            LastSentDate = $null
            LastReadMessages = @()
            Permissions = @{ FullAccess = @(); SendAs = @(); SendOnBehalf = @() }
            Error = $_.Exception.Message
        }
        
        $results += $result
    }
}

# Step 4: Generate reports
Write-Host "`nGenerating reports..." -ForegroundColor Yellow

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = Split-Path $csvFile -Parent
$htmlReportPath = Join-Path $outputDir "MailboxAnalysis_$timestamp.html"
$csvReportPath = Join-Path $outputDir "MailboxAnalysis_$timestamp.csv"

New-HTMLReport -Results $results -OutputPath $htmlReportPath
New-CSVReport -Results $results -OutputPath $csvReportPath

# Step 5: Summary
Write-Host "`n=== Analysis Complete ===" -ForegroundColor Magenta
Write-Host "Total mailboxes processed: $($results.Count)" -ForegroundColor White
Write-Host "Successful: $(($results | Where-Object { -not $_.Error }).Count)" -ForegroundColor Green
Write-Host "Failed: $(($results | Where-Object { $_.Error }).Count)" -ForegroundColor Red
Write-Host "`nReports saved to:" -ForegroundColor White
Write-Host "  HTML: $htmlReportPath" -ForegroundColor Cyan
Write-Host "  CSV:  $csvReportPath" -ForegroundColor Cyan

Write-Host "`nScript execution completed!" -ForegroundColor Green
