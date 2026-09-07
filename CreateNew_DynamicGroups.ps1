Connect-MgGraph -Scopes "Group.ReadWrite.All" -NoWelcome

try {
    $DisplayName = Read-Host "Display Name"
    $MailNickname = Read-Host "Mail Nickname"
    $Description = Read-Host "Description"
    $MembershipRule = Read-Host "Membership Rule (e.g., (user.jobTitle -eq `"Technician`"))"

    New-MgGroup `
        -DisplayName $DisplayName `
        -MailEnabled:$false `
        -MailNickname $MailNickname `
        -SecurityEnabled:$true `
        -Description $Description `
        -GroupTypes @("DynamicMembership") `
        -MembershipRule $MembershipRule `
        -MembershipRuleProcessingState "On" `
        -ErrorAction Stop

    Write-Host "Security Group '$DisplayName' created successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to create the group:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    Disconnect-MgGraph | Out-Null
}
