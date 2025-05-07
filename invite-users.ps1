$users = import-csv invitations.csv

Connect-MgGraph -Scopes "User.Invite.All"


$redirectUrl = "https://ocr.uc.edu"


foreach ($user in $users) {
    try {
        # Step 1: Check if user already exists
        $existingUser = Get-MgUser -Filter "Mail eq '$($user.Email)'" -ConsistencyLevel eventual -CountVariable count

        if ($existingUser) {
            $guest = $existingUser[0]
            if ($guest.UserType -eq "Guest") {
                Write-Host "⏩ Skipping: $($user.FullName) already invited."

                # Optional: Show invite status
                $state = $guest.ExternalUserState
                $when = $guest.ExternalUserStateChangeDateTime
                Write-Host "  📥 Invite State: $state"
                if ($state -eq "Accepted") {
                    Write-Host "  ✅ Accepted on: $when"
                } elseif ($state -eq "PendingAcceptance") {
                    Write-Host "  🕒 Pending since: $($guest.CreatedDateTime)"
                }

                continue
            } else {
                Write-Host "⚠️ $($user.FullName) exists. Skipping." -ForegroundColor Yellow
                continue
            }
        }

        # Step 2: Send invite
        $invite = New-MgInvitation -InvitedUserEmailAddress $user.Email `
                                   -InvitedUserDisplayName $user.FullName `
                                   -InviteRedirectUrl $redirectUrl `
                                   -SendInvitationMessage:$true `
                                   -InvitedUserMessageInfo @{
                                       CustomizedMessageBody = "Hello $($user.FirstName), Please click the link below to complete registration. This invite will expire in 3 weeks if not accepted. Once accepted you while be able to login to the OCR. When a class roster is submitted by your professor, you will be granted access to that course in the range.
 
If this invite has expired, please contact your professor to resubmit an access request."
                                   }

        Write-Host "✔️ Invited: $($user.FullName)"

        # Step 3: Update names
        $invitedUserId = $invite.InvitedUser.Id
        if ($invitedUserId) {
            Update-MgUser -UserId $invitedUserId -GivenName $user.FirstName -Surname $user.LastName
            Write-Host "🔄 Updated name: $($user.FirstName) $($user.LastName)"
        }
    }
    catch {
        Write-Host "❌ Error for $($user.Email): $($_.Exception.Message)" -ForegroundColor Red
    }
}

