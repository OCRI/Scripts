#TODO 
#Check to make sure latest version of powershell is installed
#Delete Completed XLSX after converstion to csv or store them somewhere
#Validate that there is data in columns
#Move invalid CSVs to rejected folder

# Define folders
$requiredColumns = @("L_NAME", "F_NAME", "EMAIL")
$rejectedFolder = "C:\Users\Public\Documents\Class Rosters\2025\Rejected"
$inputFolder = "C:\Users\Public\Documents\Class Rosters\2025\Inbound"
$processedFolder = "C:\Users\Public\Documents\Class Rosters\2025\Complete"
$historicalFolder = "C:\Users\Public\Documents\Class Rosters\2025\Historical"

# Ensure processed folder exists
if (-not (Test-Path -Path $processedFolder)) {
    Write-Host "Processed folder doesn't exist you donkey"
}

# Check if the Microsoft.Graph module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft.Graph module not found. Installing..."
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}


Connect-MgGraph -Scopes "User.Invite.All"

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
    Write-Host "Converted $($_.Name) to CSV."
    #Move orginal XLSX file
    Move-Item -Path $_.FullName -Destination (Join-Path $HistoricalFolder $_.Name)
    # Load CSV and rename column
    $csvContent = Import-Csv $csvPath
    $renamedContent = $csvContent | Select-Object @{Name='EMAIL'; Expression={ $_.'EMAIL ADDRESS for ACCESS (required)' }}, * -ExcludeProperty 'EMAIL ADDRESS for ACCESS (required)'
    $updatedData = $renamedContent | ForEach-Object {
        # Split the "LAST NAME, FIRST NAME" column
        $nameParts = $_.'LAST NAME, FIRST NAME' -split ',\s*'
        $_ | Add-Member -NotePropertyName 'L_NAME' -NotePropertyValue $nameParts[0] -Force
        $_ | Add-Member -NotePropertyName 'F_NAME' -NotePropertyValue $nameParts[1] -Force
        $_
    }

    # Save back to CSV
    $updatedData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Renamed header in: $csvPath"

}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()




$redirectUrl = "https://ocr.uc.edu"


# Process each CSV file in the input folder
Get-ChildItem -Path $inputFolder -Filter *.csv | ForEach-Object {
    $csvFile = $_.FullName
    $fileName = $_.Name
    Write-Host "Processing $fileName..."

    $users = Import-Csv $csvFile

    # Validate required columns
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $users[0].PSObject.Properties.Name }
    if ($missingColumns.Count -gt 0) {
        Write-Host "Missing columns in $fileName $($missingColumns -join ', ')" -ForegroundColor Red
        Move-Item -Path $csvFile -Destination (Join-Path $rejectedFolder $fileName)
        Write-Host "Moved to rejected folder: $fileName" -ForegroundColor DarkYellow
        return  # Skip further processing for this file
    }

    foreach ($user in $users) {
        try {
            # Step 1: Check if user already exists
            $existingUser = Get-MgUser -Filter "Mail eq '$($user.EMAIL)'" -ConsistencyLevel eventual -CountVariable count
            if ($existingUser) {
                Write-Host "$($user.F_NAME) $($user.L_NAME) exists. Skipping." -ForegroundColor Yellow
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

            Write-Host "Invited: $($user.EMAIL)"

            # Step 3: Update names
            $invitedUserId = $invite.InvitedUser.Id
            if ($invitedUserId) {
                Update-MgUser -UserId $invitedUserId -GivenName $user.F_NAME -Surname $user.L_NAME
                Write-Host "Updated name: $($user.F_NAME) $($user.L_NAME)"
            }
        }
        catch {
            Write-Host "Error for $($user.EMAIL): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Move processed file
    Move-Item -Path $csvFile -Destination (Join-Path $processedFolder $fileName)
    Write-Host "Moved $fileName to processed folder." -ForegroundColor Cyan
}

Write-Host "`nAll files processed." -ForegroundColor Green