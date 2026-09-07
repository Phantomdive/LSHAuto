function Set-CalendarPermissions {

    Connect-ExchangeOnline -ShowBanner:$false

    $mailbox = Read-Host "Mailbox Email Address"

    do {

        Clear-Host

        Write-Host "========================================"
        Write-Host " Calendar Permission Management"
        Write-Host " Mailbox: $mailbox"
        Write-Host "========================================"
        Write-Host ""
        Write-Host "1. View Permissions"
        Write-Host "2. Grant Reviewer Access"
        Write-Host "3. Grant Editor Access"
        Write-Host "4. Grant Owner Access"
        Write-Host "5. Grant AvailabilityOnly Access"
        Write-Host "6. Remove Access"
        Write-Host "7. Exit"
        Write-Host ""

        $choice = Read-Host "Select option"

        switch ($choice) {

            '1' {
                Get-MailboxFolderPermission `
                    -Identity "${mailbox}:\Calendar" |
                    Select User,AccessRights |
                    Format-Table -AutoSize

                Pause
            }

            '2' {
                $user = Read-Host "User Email Address"

                try {
                    Set-MailboxFolderPermission `
                        -Identity "${mailbox}:\Calendar" `
                        -User $user `
                        -AccessRights Reviewer `
                        -ErrorAction Stop
                }
                catch {
                    Add-MailboxFolderPermission `
                        -Identity "${mailbox}:\Calendar" `
                        -User $user `
                        -AccessRights Reviewer
                }

                Write-Host "Reviewer access granted." -ForegroundColor Green
                Pause
            }

            '3' {
                $user = Read-Host "User Email Address"

                try {
                    Set-MailboxFolderPermission `
                        -Identity "${mailbox}:\Calendar" `
                        -User $user `
                        -AccessRights Editor `
                        -ErrorAction Stop
                }
                catch {
                    Add-MailboxFolderPermission `
                        -Identity "${mailbox}:\Calendar" `
                        -User $user `
                        -AccessRights Editor
                }

                Write-Host "Editor access granted." -ForegroundColor Green
                Pause
            }

            '4' {
                $user = Read-Host "User Email Address"

                try {
                    Set-MailboxFolderPermission `
                        -Identity "${mailbox}:\Calendar" `
                        -User $user `
                        -AccessRights Owner `
                        -ErrorAction Stop
                }
                catch {
                    Add-MailboxFolderPermission `
                        -Identity "${mailbox}:\Calendar" `
                        -User $user `
                        -AccessRights Owner
                }

                Write-Host "Owner access granted." -ForegroundColor Green
                Pause
            }

            '5' {
                $user = Read-Host "User Email Address"

                try {
                    Set-MailboxFolderPermission `
                        -Identity "${mailbox}:\Calendar" `
                        -User $user `
                        -AccessRights AvailabilityOnly `
                        -ErrorAction Stop
                }
                catch {
                    Add-MailboxFolderPermission `
                        -Identity "${mailbox}:\Calendar" `
                        -User $user `
                        -AccessRights AvailabilityOnly
                }

                Write-Host "AvailabilityOnly access granted." -ForegroundColor Green
                Pause
            }


            '6' {
                $user = Read-Host "User Email Address"

                Remove-MailboxFolderPermission `
                    -Identity "${mailbox}:\Calendar" `
                    -User $user `
                    -Confirm:$false

                Write-Host "Permission removed." -ForegroundColor Green
                Pause
            }
            
            '7' {
                Disconnect-ExchangeOnline -Confirm:$false
                break
            }

            default {
                Write-Host "Invalid selection." -ForegroundColor Red
                Pause
            }
        }

    } while ($true)
}

Set-CalendarPermissions