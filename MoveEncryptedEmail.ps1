<#
.SYNOPSIS
Move RMS protected emails for multiple mailboxes to a specific folder.
With dynamic mailbox list (from Git repo) and logs stored in GCS.
#>

# ---------------- CONFIGURATION ----------------
$TenantId       = $env:TENANT_ID
$ClientId       = $env:CLIENT_ID
$ClientSecret   = $env:CLIENT_SECRET    # stored in Secret Manager or env var
$RepoUrl        = "https://github.com/mathursakshamonix/encrypted_mailbox.git"
$RepoPath       = "."
$TargetFolder   = "RMSAIPEncryptedEmails"
$MonthsBack     = 24
$LogFile        = "/logs/newencryptedmoved_emails.csv"
$GcsBucket      = $env:GCS_BUCKET       # e.g. my-mailbox-logs

"UserID,Sender,Recipients,ItemID,ReceivedDate,Subject,MessageID" | Out-File -FilePath $LogFile -Encoding UTF8

# ---------------- FETCH MAILBOX CSV FROM GIT ----------------
if (Test-Path $RepoPath) {
    git -C $RepoPath pull
} else {
    git clone $RepoUrl $RepoPath
}
$CsvPath = "$RepoPath/mailbox.csv"
Write-Host "Using CSV: $CsvPath"

# ---------------- AUTHENTICATION ----------------
$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$clientSecretCredential = New-Object Microsoft.Graph.PowerShell.Authentication.Core.ClientSecretCredential `
    ($TenantId, $ClientId, $secureSecret)

Connect-MgGraph -ClientSecretCredential $clientSecretCredential -TenantId $TenantId
Write-Host "Connected to Microsoft Graph" -ForegroundColor Green

# ---------------- FUNCTIONS ----------------
function Invoke-WithRetry {
    param ([ScriptBlock]$Script, [int]$MaxRetries = 5)
    $retry = 0; $delay = 2
    while ($true) {
        try { return & $Script }
        catch {
            $errorMsg = $_.Exception.Message
            if ($_ -match "StatusCode: 429" -or $_ -match "StatusCode: 503") {
                if ($retry -ge $MaxRetries) {
                    Write-Host "Max retries reached. Skipping..." -ForegroundColor Yellow
                    return $null
                }
                $retry++; Write-Host "Retry $retry after $delay sec" -ForegroundColor Yellow
                Start-Sleep -Seconds $delay
                $delay = [math]::Min($delay * 2, 60)
            } else {
                Write-Host "ERROR: $errorMsg" -ForegroundColor Red
                return $null
            }
        }
    }
}

function Get-OrCreateFolder {
    param ([string]$UserId, [string]$FolderName)
    $folder = Invoke-WithRetry { Get-MgUserMailFolder -UserId $UserId -All | Where-Object { $_.DisplayName -eq $FolderName } }
    if (-not $folder) {
        Write-Host "Creating folder $FolderName for $UserId"
        $folder = Invoke-WithRetry { New-MgUserMailFolder -UserId $UserId -DisplayName $FolderName }
    }
    return $folder.Id
}

function Test-AIPAttachment {
    param([object]$Message, [string]$UserId)
    if (-not $Message.HasAttachments) { return $false }
    try {
        $attachments = Invoke-WithRetry { Get-MgUserMessageAttachment -UserId $UserId -MessageId $Message.Id -All }
        foreach ($att in $attachments) {
            if ($att.AdditionalProperties.contentBytes) {
                $bytes   = [System.Convert]::FromBase64String($att.AdditionalProperties.contentBytes)
                $content = [System.Text.Encoding]::UTF8.GetString($bytes)
                if ($content -match "DRMEncryptedTransform|Label ID|EncryptedPackage|RightsManagement") {
                    return $true
                }
            }
        }
    } catch { Write-Warning "Attachment fetch failed for $($Message.Id)" }
    return $false
}

# ---------------- MAIN ----------------
$startDateTime = (Get-Date).AddMonths(-$MonthsBack).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$mailboxes = Import-Csv -Path $CsvPath

foreach ($mb in $mailboxes) {
    $userId = $mb.UserPrincipalName
    Write-Host "Processing mailbox: $userId" -ForegroundColor Cyan

    $targetFolderId = Get-OrCreateFolder -UserId $userId -FolderName $TargetFolder
    if (-not $targetFolderId) { continue }

    $folders = Invoke-WithRetry { Get-MgUserMailFolder -UserId $userId -All }
    foreach ($folder in $folders) {
        $nextLink = $null
        do {
            $messagesResponse = if ($nextLink) {
                Invoke-MgGraphRequest -Uri $nextLink -Method GET
            } else {
                Get-MgUserMailFolderMessage -UserId $userId -MailFolderId $folder.Id `
                    -Filter "receivedDateTime ge $startDateTime" `
                    -Property "InternetMessageHeaders,Subject,receivedDateTime,from,toRecipients,id,HasAttachments" `
                    -All -PageSize 50
            }
            if (-not $messagesResponse) { break }
            $messages = if ($messagesResponse.value) { $messagesResponse.value } else { $messagesResponse }

            foreach ($message in $messages) {
                $isEncrypted = $message.InternetMessageHeaders | Where-Object { $_.Name -ieq "Microsoft.Exchange.RMSApaAgent.ProtectionTemplateId" }
                $hasAIPAttachment = if (-not $isEncrypted) { Test-AIPAttachment -Message $message -UserId $userId } else { $false }

                if ($isEncrypted -or $hasAIPAttachment) {
                    Write-Host "$userId: Found encrypted mail → $($message.Subject)"
                    $sender    = $message.From.EmailAddress.Address
                    $recipients = ($message.ToRecipients | ForEach-Object { $_.EmailAddress.Address }) -join ";"
                    $msgId     = $message.Id
                    $date      = $message.ReceivedDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                    $subject   = $message.Subject -replace '"', '""'
                    $messageId = ($message.InternetMessageHeaders | Where-Object { $_.Name -eq "Message-ID" }).Value

                    "$userId,$sender,$recipients,$msgId,$date,$subject,$messageId" | Out-File -FilePath $LogFile -Append -Encoding UTF8
                    Move-MgUserMessage -UserId $userId -MessageId $message.Id -DestinationId $targetFolderId | Out-Null
                }
            }
            $nextLink = $messagesResponse.'@odata.nextLink'
        } while ($nextLink)
    }
}

Write-Host "Completed processing all mailboxes." -ForegroundColor Green

# ---------------- UPLOAD LOGS TO GCS ----------------
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$dest = "gs://$GcsBucket/logs/encrypted_moved_$timestamp.csv"
Write-Host "Uploading logs to $dest"
& gsutil cp $LogFile $dest
