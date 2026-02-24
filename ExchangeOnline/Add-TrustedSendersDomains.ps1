[CmdletBinding(PositionalBinding=$false)]
param(
    [string[]]$locations = @(),
    [string]$locationsPath = "/Users/jamiell/Developer/Scripts/locations.txt",
    [string]$domainsPath = "/Users/jamiell/Developer/Scripts/domains.txt",
    [string]$errorlog = "/Users/jamiell/Developer/Scripts/errors.csv",
    [boolean]$autoDisconnectEXO = $false,
    [boolean]$recycleMailboxes = $true
)

$totalfailures = 0;
$totalSuccesses = 0;

if(($null -ne $locationsPath) -and (Test-Path $locationsPath) -and ($locations.Count -le 0)) {
    $locations = Get-Content -Path $locationsPath | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    if($locations.Count -le 0) {
        Write-Host "No locations have been found in the file. Please check the file and try again." -ForegroundColor Yellow
        exit
    }
}

if(($null -ne $domainsPath) -and (Test-Path -Path $domainsPath)) {
    $domains = Get-Content -Path $domainsPath | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    if($domains.Count -le 0) {
        Write-Host "No domains have been found in the file. Please check the file and try again." -ForegroundColor Yellow
        exit
    }
}

if(Get-ConnectionInformation)
{
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
}
else
{
    Write-Host "Not connected to Exchange Online. Attempting to connect..." -ForegroundColor Yellow
    Connect-ExchangeOnline 
}

Write-Host "Fetching user mailboxes..." -ForegroundColor Cyan

$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Select-Object -Property UserPrincipalName, Office

if($mailboxes.Count -le 0) {
    Write-Host "No mailboxes have been found. Please check your connection and try again." -ForegroundColor Yellow
    exit
}

foreach ($location in $locations) {
    Write-Host "`nFetching mailboxes for location: $location..." -ForegroundColor Cyan
    $filtered = @($mailboxes | Where-Object { $_.Office -eq $location})

    if($filtered.Count -le 0) {
        Write-Host "`nNo mailboxes found for location: $location" -ForegroundColor Yellow
        continue
    }

    $count = 0;
    $successes = 0
    $failures = 0

    foreach ($mailbox in $filtered) {
        $count++
        try {
            Set-MailboxJunkEmailConfiguration -Identity $mailbox.UserPrincipalName -TrustedSendersAndDomains @{Add=$domains} -ErrorAction Stop
            $successes++ 
            $totalSuccesses++;
        }
        catch {
            #Write-Host "Failed to update mailbox: $($mailbox.UserPrincipalName). Error: $_" -ForegroundColor Red
            $errors = @{
                Mailbox = $mailbox.UserPrincipalName;
                Location = $location;
                Error = $_.Exception.Message
            }

            $failures++
            $totalfailures++;

            $errors | Export-Csv -Path $errorlog -Append -Force
        }
        Write-Host "`rProcessing mailbox ($count / $($filtered.Count)) | Location: $location | Successful: $successes | Failed: $failures" -NoNewline
    }
}


if($autoDisconnectEXO) 
{
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "`nDisconnected from Exchange Online." -ForegroundColor Green
}

Write-Host "`nScript execution completed. Check $errorlog for any errors." -ForegroundColor Green
