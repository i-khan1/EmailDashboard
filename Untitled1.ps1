# EmailDashboard.ps1

Add-Type -AssemblyName Microsoft.Office.Interop.Outlook

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)
$folder = $inbox.Folders.Item("reports")

$targetSender = "reports@velocityclearingllc.com"
$today = Get-Date
$todayDate = $today.ToString("yyyyMMdd")

# Collect today's emails
$items = $folder.Items
$items.Sort("[ReceivedTime]", $false)
$emailsToday = @()

foreach ($mail in $items) {
    if ($mail.Class -eq 43 -and $mail.ReceivedTime -and $mail.SenderEmailAddress -eq $targetSender) {
        $mailDate = ($mail.ReceivedTime -as [datetime]).ToString("yyyyMMdd")
        if ($mailDate -eq $todayDate) {
            $emailsToday += [PSCustomObject]@{
                Subject      = $mail.Subject
                ExpectedTime = ""
                ReceivedTime = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                From         = $mail.SenderEmailAddress
                To           = $mail.To
                CC           = $mail.CC
                BCC          = $mail.BCC
                Status       = "Present"
            }
        }
    }
}

# Save new data temporarily
$tempJson = "D:\EmailDashboard\temp_email.json"
$emailsToday | ConvertTo-Json -Depth 3 | Set-Content $tempJson -Encoding UTF8

# Compare hashes
$jsonPath = "D:\EmailDashboard\emails.json"
$oldHash = if (Test-Path $jsonPath) { Get-FileHash $jsonPath } else { $null }
$newHash = Get-FileHash $tempJson

if (-not $oldHash -or $oldHash.Hash -ne $newHash.Hash) {
    # Data changed — update real file and push
    Copy-Item $tempJson $jsonPath -Force
    Set-Location "D:\EmailDashboard"

    git add email.json
    git commit -m "Auto update email data $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" 2>$null
    git push origin main
} else {
    Write-Output "No change in email data. Skipping push."
}
