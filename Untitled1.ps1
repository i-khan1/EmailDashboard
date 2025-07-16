Add-Type -AssemblyName Microsoft.Office.Interop.Outlook

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$inbox = $namespace.GetDefaultFolder(6)
$folder = $inbox.Folders.Item("reports")

$targetSender = "reports@velocityclearingllc.com"
$todayDate = Get-Date -Format "yyyyMMdd"
$today = Get-Date

# --- Load Today's Emails from Outlook ---
$items = $folder.Items
$items.Sort("[ReceivedTime]", $false)
$emailsToday = @()

foreach ($mail in $items) {
    if ($mail.Class -eq 43 -and $mail.ReceivedTime -and $mail.SenderEmailAddress -eq $targetSender) {
        $mailDate = ($mail.ReceivedTime -as [datetime]).ToString("yyyyMMdd")
        if ($mailDate -eq $todayDate) {
            $emailsToday += $mail
        }
    }
}

# --- Step 1: Maintain Historical Log ---
$historyPath = "C:\Users\Ilma Khan\email-monitor\backend\history_log.json"

if (Test-Path $historyPath) {
    $history = Get-Content $historyPath | ConvertFrom-Json
} else {
    $history = @()
}

# Prevent duplication in history
foreach ($mail in $emailsToday) {
    $entryDate = $mail.ReceivedTime.ToString("yyyy-MM-dd")
    $entryTime = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
    $exists = $history | Where-Object { $_.Subject -eq $mail.Subject -and $_.ReceivedTime -eq $entryTime }

    if (-not $exists) {
        $history += [PSCustomObject]@{
            Subject      = $mail.Subject
            ReceivedTime = $entryTime
            Date         = $entryDate
        }
    }
}

# Trim history to only last 7 days
$cutoff = (Get-Date).AddDays(-7)
$history = $history | Where-Object { [datetime]$_.Date -ge $cutoff }

# Save updated history
$history | ConvertTo-Json -Depth 3 | Set-Content $historyPath

# --- Step 2: Learn Expected Times Based on History (7 days) ---
$recentHistory = $history

$learnedSchedule = $recentHistory | Group-Object Subject | ForEach-Object {
    $subject = $_.Name
    $times = $_.Group | ForEach-Object { [datetime]$_.ReceivedTime }
    $timeBuckets = $times | Group-Object { $_.ToString("HH:mm") } | Sort-Object Count -Descending
    $mostFrequentTime = $timeBuckets[0].Name

    [PSCustomObject]@{
        Subject      = $subject
        ExpectedTime = $mostFrequentTime
    }
}

# --- Step 3: Check Which Subjects Are Missing Today ---
$reportStatus = @()

foreach ($entry in $learnedSchedule) {
    $expectedDateTime = [datetime]::ParseExact("$($today.ToString("yyyy-MM-dd")) $($entry.ExpectedTime)", "yyyy-MM-dd HH:mm", $null)

    $matchedEmail = $emailsToday | Where-Object {
        $_.Subject -eq $entry.Subject -and
        ($_.ReceivedTime -ge $expectedDateTime.AddMinutes(-5)) -and
        ($_.ReceivedTime -le $expectedDateTime.AddMinutes(5))
    } | Select-Object -First 1

    if ($matchedEmail) {
        $reportStatus += [PSCustomObject]@{
            Subject      = $entry.Subject
            ExpectedTime = $expectedDateTime.ToString("yyyy-MM-dd HH:mm")
            ReceivedTime = $matchedEmail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
            From         = $matchedEmail.SenderEmailAddress
            To           = $matchedEmail.To
            CC           = $matchedEmail.CC
            BCC          = $matchedEmail.BCC
            Status       = "Present"
        }
    } else {
        $reportStatus += [PSCustomObject]@{
            Subject      = $entry.Subject
            ExpectedTime = $expectedDateTime.ToString("yyyy-MM-dd HH:mm")
            ReceivedTime = $null
            From         = $targetSender
            To           = ""
            CC           = ""
            BCC          = ""
            Status       = "Missing"
        }
    }
}

# --- Step 4: Export Final Report ---
$reportPath = "C:\Users\Ilma Khan\email-monitor\backend\emails.json"
$reportStatus | ConvertTo-Json -Depth 3 | Set-Content $reportPath
