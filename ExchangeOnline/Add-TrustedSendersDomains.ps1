param(
    [string]$location = "Jersey City",
    [string]$path = "/Users/jamiell/Developer/Scripts/domains.txt",
    [string]$errorlog = "/Users/jamiell/Developer/Scripts/errors_$location.csv",
    [boolean]$autoDisconnectEXO = $false,
    [boolean]$recycleMailboxes = $true
)


$successes = 0
$failures = @()

$domains = Get-Content -Path $path | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

if(Get-ConnectionInformation)
{
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
}
else
{
    Write-Host "Not connected to Exchange Online. Attempting to connect..." -ForegroundColor Yellow
    Connect-ExchangeOnline 
}

Write-Host "Fetching mailboxes for location: $location..." -ForegroundColor Cyan
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Select-Object -Property UserPrincipalName, Office | Where-Object { $_.Office -eq $location}

if($mailboxes.Count -eq 0) {
    Write-Host "No mailboxes found in $location."
    exit
}

$count = 0;
foreach ($mailbox in $mailboxes) {
    count++;
    Write-Host "`rProcessing mailbox ($count / $($mailboxes.Count)) | Location: $location | Successful: $successes | Failed: $($failures.Count)" -NoNewline
    try {
        Set-MailboxJunkEmailConfiguration -Identity $mailbox.UserPrincipalName -TrustedSendersAndDomains @{Add=$domains} -ErrorAction Stop
        $successes++

    }
    catch {
        #Write-Host "Failed to update mailbox: $($mailbox.UserPrincipalName). Error: $_" -ForegroundColor Red
        $failures += @{
            Mailbox = $mailbox.UserPrincipalName;
            Location = $location;
            Error = $_.Exception.Message
        }
    }
}

if($autoDisconnectEXO) 
{
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "`nDisconnected from Exchange Online." -ForegroundColor Green
}

if($failures.Count -le 0) 
{
    Write-Host "`nAll mailboxes updated successfully." -ForegroundColor Green
}
else 
{
    Write-Host "`nSome mailboxes failed to update. See failures at $errorlog." -ForegroundColor Red
    $errorlog = $errorlog.Trim() -replace '[\[\]]', ''
    $failures | Export-Csv -Path $errorlog -Force
}