# Define folders
$inputFolder = "C:\Users\Public\Documents\Class Rosters\2025\Inbound"
$processedFolder = "C:\Users\Public\Documents\Class Rosters\2025\Complete"

# Requires Excel COM object (Excel must be installed)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Convert all .xlsx files to .csv
Get-ChildItem -Path $inputFolder -Filter *.xlsx | ForEach-Object {
    $workbook = $excel.Workbooks.Open($_.FullName)
    $csvPath = ($_.FullName -replace '\.xlsx$', '.csv')
    $workbook.SaveAs($csvPath, 6)  # 6 = xlCSV
    $workbook.Close($false)
    Write-Host "📝 Converted $($_.Name) to CSV."
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# Ensure processed folder exists
if (-not (Test-Path -Path $processedFolder)) {
    Write-Host "Processed folder doesn't exist you donkey"
}

# Check if the Microsoft.Graph module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft.Graph module not found. Installing..."
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}

# Import the module (just in case it's not auto-loaded)
Import-Module Microsoft.Graph

Connect-MgGraph -Scopes "User.Invite.All"


$redirectUrl = "https://ocr.uc.edu"


# Process each CSV file in the input folder
Get-ChildItem -Path $inputFolder -Filter *.csv | ForEach-Object {
    $csvFile = $_.FullName
    $fileName = $_.Name
    Write-Host "📄 Processing $fileName..."

    $users = Import-Csv $csvFile

    foreach ($user in $users) {
        try {
            # Step 1: Check if user already exists
            $existingUser = Get-MgUser -Filter "Mail eq '$($user.EMAIL)'" -ConsistencyLevel eventual -CountVariable count
            if ($existingUser) {
                Write-Host "⚠️ $($user.F_NAME) $($user.L_NAME) exists. Skipping." -ForegroundColor Yellow
                continue
            }

            # Step 2: Send invite
            $invite = New-MgInvitation -InvitedUserEmailAddress $user.EMAIL `
                                       -InvitedUserDisplayName "$($user.F_NAME) $($user.L_NAME)" `
                                       -InviteRedirectUrl $redirectUrl `
                                       -SendInvitationMessage:$true `
                                       -InvitedUserMessageInfo @{
                                           CustomizedMessageBody = "Hello $($user.F_NAME), Please click the link below to complete registration. This invite will expire in 3 weeks if not accepted. Once accepted you will be able to login to the OCR. When a class roster is submitted by your professor, you will be granted access to that course in the range.

If this invite has expired, please contact your professor to resubmit an access request."
                                       }

            Write-Host "✔️ Invited: $($user.EMAIL)"

            # Step 3: Update names
            $invitedUserId = $invite.InvitedUser.Id
            if ($invitedUserId) {
                Update-MgUser -UserId $invitedUserId -GivenName $user.F_NAME -Surname $user.L_NAME
                Write-Host "🔄 Updated name: $($user.F_NAME) $($user.L_NAME)"
            }
        }
        catch {
            Write-Host "❌ Error for $($user.EMAIL): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Move processed file
    Move-Item -Path $csvFile -Destination (Join-Path $processedFolder $fileName)
    Write-Host "📁 Moved $fileName to processed folder." -ForegroundColor Cyan
}

Write-Host "`n✅ All files processed." -ForegroundColor Green