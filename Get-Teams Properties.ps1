Connect-ExchangeOnline
Connect-MicrosoftTeams
# Get the properties of a specific Teams group
$teamName = 'Marketing CXC 2022'
$team = Get-Team | Where-Object { $_.DisplayName -eq $teamName }

# Display the properties of the team
$team | Format-List

# Get the Unified Group and retrieve the expiration date
$unifiedGroup = Get-UnifiedGroup | Where-Object { $_.DisplayName -eq $teamName }

# Display the expiration date
if ($unifiedGroup) {
    $expirationDate = $unifiedGroup.ExpirationTime
    Write-Output "The expiration date for the group '$teamName' is: $expirationDate"
} else {
    Write-Output "Group '$teamName' not found."
}
#







# Connect to Exchange Online
#Connect-ExchangeOnline

# Retrieve all Unified Groups and their expiration dates
#$allGroups = Get-UnifiedGroup

# Display the expiration date for each group
#foreach ($group in $allGroups) {
#    $groupName = $group.DisplayName
#    $expirationDate = $group.ExpirationTime

#    if ($expirationDate) {
 #       Write-Output "The expiration date for the group '$groupName' is: $expirationDate"
 #  } else {
#     Write-Output "The group '$groupName' does not have an expiration date set."
    #}
#}