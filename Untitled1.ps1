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
                ExpectedTime = ""  # Optional placeholder
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

# Save to JSON
$jsonPath = "D:\EmailDashboard\emails.json"
$emailsToday | ConvertTo-Json -Depth 3 | Set-Content $jsonPath -Encoding UTF8

# Change directory to Git repo
Set-Location "D:\EmailDashboard"

# Git commit and push
git add email.json
git commit -m "Auto update email data $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" 2>$null
git push origin main
