<#
.SYNOPSIS
    Creates Microsoft Entra ID users through interactive or CSV entry points.

.DESCRIPTION
    Supports creating one user interactively or multiple users from a CSV file.

    The script can:
    - Create Member or Guest user objects
    - Populate user profile information
    - Assign a manager
    - Assign an E3, E5 or F3 licence
    - Generate a secure temporary password
    - Support -WhatIf and -Confirm
#>


[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Menu', 'Single', 'Csv')]
    [string]$EntryPoint = 'Menu',

    [string]$CsvPath,

    [switch]$EnableUser,

    [switch]$ForcePasswordChange
)

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Force -AllowClobber -Scope AllUsers
}

$ErrorActionPreference = 'Stop'

$GraphScopes = @(
    'User.ReadWrite.All'
    'Directory.ReadWrite.All'
    'Organization.Read.All'
    'LicenseAssignment.ReadWrite.All'
)

function Connect-Entra {
    [CmdletBinding()]
    param()

    $requiredModules = @(
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Users.Actions'
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw @"
Required Microsoft Graph module '$module' is not installed.

Install the full Microsoft Graph PowerShell SDK using:

Install-Module Microsoft.Graph -Scope CurrentUser
"@
        }
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    $context = Get-MgContext

    $mustReconnect = $false

    if (-not $context) {
        $mustReconnect = $true
    }
    else {
        $currentScopes = @($context.Scopes)

        foreach ($requiredScope in $GraphScopes) {
            if ($requiredScope -notin $currentScopes) {
                $mustReconnect = $true
                break
            }
        }
    }

    if ($mustReconnect) {
        Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan

        Connect-MgGraph `
            -Scopes $GraphScopes `
            -NoWelcome `
            -ContextScope Process
    }

    $context = Get-MgContext

    if (-not $context) {
        throw 'Microsoft Graph connection was not established.'
    }

    Write-Host "Connected to Microsoft Graph as $($context.Account)." -ForegroundColor Green
}

function Read-Required {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    do {
        $value = Read-Host $Prompt

        if ($null -ne $value) {
            $value = $value.Trim()
        }
    }
    while ([string]::IsNullOrWhiteSpace($value))

    return $value
}

function ConvertTo-CleanAlias {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Alias
    )

    $cleanAlias = $Alias.Trim().ToLowerInvariant()

    # Replace spaces with dots.
    $cleanAlias = $cleanAlias -replace '\s+', '.'

    # Remove characters that should not appear in the alias or UPN.
    $cleanAlias = $cleanAlias -replace '[^a-z0-9._-]', ''

    # Remove duplicate dots.
    $cleanAlias = $cleanAlias -replace '\.{2,}', '.'

    # Remove dots, underscores and hyphens from the beginning and end.
    $cleanAlias = $cleanAlias.Trim('.', '_', '-')

    if ([string]::IsNullOrWhiteSpace($cleanAlias)) {
        throw 'The supplied username does not contain any valid characters.'
    }

    return $cleanAlias
}

function New-TemporaryPassword {
    [CmdletBinding()]
    param(
        [ValidateRange(12, 128)]
        [int]$Length = 18
    )

    $upperCase = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $lowerCase = 'abcdefghijkmnopqrstuvwxyz'
    $numbers   = '23456789'
    $symbols   = '!@#$%*-_=+?'
    $allChars  = $upperCase + $lowerCase + $numbers + $symbols

    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()

    try {
        function Get-CryptoRandomIndex {
            param(
                [Parameter(Mandatory)]
                [int]$Maximum
            )

            $bytes = New-Object byte[] 4
            $rng.GetBytes($bytes)

            $unsignedNumber = [BitConverter]::ToUInt32($bytes, 0)

            return $unsignedNumber % [uint32]$Maximum
        }

        $passwordCharacters = New-Object 'System.Collections.Generic.List[char]'

        # Guarantee at least one character from each category.
        $passwordCharacters.Add(
            $upperCase[(Get-CryptoRandomIndex -Maximum $upperCase.Length)]
        )

        $passwordCharacters.Add(
            $lowerCase[(Get-CryptoRandomIndex -Maximum $lowerCase.Length)]
        )

        $passwordCharacters.Add(
            $numbers[(Get-CryptoRandomIndex -Maximum $numbers.Length)]
        )

        $passwordCharacters.Add(
            $symbols[(Get-CryptoRandomIndex -Maximum $symbols.Length)]
        )

        while ($passwordCharacters.Count -lt $Length) {
            $passwordCharacters.Add(
                $allChars[(Get-CryptoRandomIndex -Maximum $allChars.Length)]
            )
        }

        # Shuffle the password so the required character types are not always
        # in the first four positions.
        for ($index = $passwordCharacters.Count - 1; $index -gt 0; $index--) {
            $swapIndex = Get-CryptoRandomIndex -Maximum ($index + 1)

            $temporaryCharacter = $passwordCharacters[$index]
            $passwordCharacters[$index] = $passwordCharacters[$swapIndex]
            $passwordCharacters[$swapIndex] = $temporaryCharacter
        }

        return -join $passwordCharacters
    }
    finally {
        $rng.Dispose()
    }
}

function Get-UserInput {
    [CmdletBinding()]
    param()

    $userTypeSelection = Read-Host 'User type: 1 - Normal user, 2 - Guest'

    switch ($userTypeSelection) {
        '1' {
            $userType = 'Member'
        }

        '2' {
            $userType = 'Guest'
        }

        default {
            throw "Invalid user type selection: '$userTypeSelection'."
        }
    }

    $siteSelection = Read-Host 'Site: 1 - lshauto.co.uk, 2 - lsh-properties.co.uk'

    switch ($siteSelection) {
        '1' {
            $domain = 'lshauto.co.uk'
            $companyname = 'LSH Auto UK Ltd'
        }

        '2' {
            $domain = 'lsh-properties.co.uk'
            $companyname = 'LSH Properties'
        }

        default {
            throw "Invalid site selection: '$siteSelection'."
        }
    }

    Write-Host ''
    Write-Host 'Available locations:' -ForegroundColor Cyan
    Write-Host '1 - Stockport'
    Write-Host '2 - Birmingham'
    Write-Host '3 - Macclesfield'
    Write-Host '4 - Whitefield'
    Write-Host '5 - Solihull'
    Write-Host '6 - Tamworth'
    Write-Host '7 - London'

    $citySelection = Read-Host 'Select a city'

    switch ($citySelection) {
        '1' {
            $cityName      = 'Stockport'
            $officeLocation = 'Brighton Road, Stockport, SK4 2BE'
            $streetAddress  = 'Brighton Road'
            $postalCode     = 'SK4 2BE'
            $state          = 'Greater Manchester'
        }

        '2' {
            $cityName       = 'Birmingham'
            $officeLocation = '144 Bromford Lane, Erdington, West Midlands, B24 8DE'
            $streetAddress  = '144 Bromford Lane'
            $postalCode     = 'B24 8DE'
            $state          = 'West Midlands'
        }

        '3' {
            $cityName       = 'Macclesfield'
            $officeLocation = 'Lyme Green Business Park, Brindley Way, Macclesfield, SK11 0TB'
            $streetAddress  = 'Lyme Green Business Park, Brindley Way'
            $postalCode     = 'SK11 0TB'
            $state          = 'Cheshire'
        }

        '4' {
            $cityName       = 'Whitefield'
            $officeLocation = '845 Manchester Road, Bury, Greater Manchester, BL9 9TP'
            $streetAddress  = '845 Manchester Road'
            $postalCode     = 'BL9 9TP'
            $state          = 'Greater Manchester'
        }

        '5' {
            $cityName       = 'Solihull'
            $officeLocation = 'Focus Park, Ashbourne Way, Solihull, West Midlands, B90 4QU'
            $streetAddress  = 'Focus Park, Ashbourne Way'
            $postalCode     = 'B90 4QU'
            $state          = 'West Midlands'
        }

        '6' {
            $cityName       = 'Tamworth'
            $officeLocation = 'Hints Road, Mile Oak, Tamworth, B78 3PQ'
            $streetAddress  = 'Hints Road, Mile Oak'
            $postalCode     = 'B78 3PQ'
            $state          = 'Staffordshire'
        }

        '7' {
            $cityName       = 'London'
            $officeLocation = 'Westfield Shopping Centre, London, W12 7GF'
            $streetAddress  = 'Westfield Shopping Centre'
            $postalCode     = 'W12 7GF'
            $state          = 'Greater London'
        }

        default {
            throw "Invalid city selection: '$citySelection'."
        }
    }

    $country       = 'United Kingdom'
    $usageLocation = 'GB'

    $givenName = Read-Required -Prompt 'First name'
    $surname   = Read-Required -Prompt 'Last name'
    $department = Read-Required -Prompt 'Department'
    $jobTitle   = Read-Required -Prompt 'Job title'
    $ExtensionAttribute1 = Read-Host 'Extension Attribute 1 (optional)'
    $ExtensionAttribute2 = Read-Host 'Extension Attribute 2 (optional)'

    $defaultAlias = "$($givenName.ToLowerInvariant()).$($surname.ToLowerInvariant())"

    $aliasInput = Read-Host "User name [$defaultAlias]"

    if ([string]::IsNullOrWhiteSpace($aliasInput)) {
        $aliasInput = $defaultAlias
    }

    $alias = ConvertTo-CleanAlias -Alias $aliasInput

    $managerEmail = Read-Host 'Manager email or UPN (optional)'

    if ($null -ne $managerEmail) {
        $managerEmail = $managerEmail.Trim()
    }

    $licenceSelection = Read-Host 'Licence: 1 - E3, 2 - E5, 3 - F3, 4 - None'

    switch ($licenceSelection) {
        '1' {
            $licenceSku = 'SPE_E3'
        }

        '2' {
            $licenceSku = 'SPE_E5'
        }

        '3' {
            $licenceSku = 'SPE_F1'
        }

        '4' {
            $licenceSku = $null
        }

        default {
            throw "Invalid licence selection: '$licenceSelection'."
        }
    }

    return [pscustomobject]@{
        DisplayName       = "$givenName $surname"
        GivenName         = $givenName
        Surname           = $surname
        MailNickname      = $alias
        UserPrincipalName = "$alias@$domain"
        UserType          = $userType
        City              = $cityName
        State             = $state
        StreetAddress     = $streetAddress
        PostalCode        = $postalCode
        Country           = $country
        OfficeLocation    = $officeLocation
        UsageLocation     = $usageLocation
        ManagerEmail      = $managerEmail
        LicenseSku        = $licenceSku
        Department        = $department
        JobTitle          = $jobTitle
        CompanyName       = $companyname
        ExtensionAttribute1 = $ExtensionAttribute1
        ExtensionAttribute2 = $ExtensionAttribute2
        Password          = $null
    }
}

function ConvertTo-UserRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [psobject]$Row
    )

    $requiredProperties = @(
        'DisplayName'
        'GivenName'
        'Surname'
        'UserPrincipalName'
    )

    foreach ($requiredProperty in $requiredProperties) {
        if ([string]::IsNullOrWhiteSpace([string]$Row.$requiredProperty)) {
            throw "Required user property '$requiredProperty' is missing."
        }
    }

    $mailNickname = $Row.MailNickname

    if ([string]::IsNullOrWhiteSpace([string]$mailNickname)) {
        $mailNickname = ($Row.UserPrincipalName -split '@')[0]
    }

    $mailNickname = ConvertTo-CleanAlias -Alias $mailNickname

    if ([string]::IsNullOrWhiteSpace([string]$Row.Password)) {
        $plainTextPassword = New-TemporaryPassword
        $passwordWasGenerated = $true
    }
    else {
        $plainTextPassword = [string]$Row.Password
        $passwordWasGenerated = $false
    }

    if ([string]::IsNullOrWhiteSpace([string]$Row.UsageLocation)) {
        $usageLocation = 'GB'
    }
    else {
        $usageLocation = ([string]$Row.UsageLocation).Trim().ToUpperInvariant()
    }

    if ([string]::IsNullOrWhiteSpace([string]$Row.UserType)) {
        $userType = 'Member'
    }
    else {
        $userType = [string]$Row.UserType
    }

    $forcePasswordChangeValue = $true

    if ($PSBoundParameters.ContainsKey('ForcePasswordChange')) {
        $forcePasswordChangeValue = [bool]$ForcePasswordChange
    }

    # For this account-creation script, generated passwords should always be
    # changed at the first sign-in.
    if ($passwordWasGenerated) {
        $forcePasswordChangeValue = $true
    }

    $request = @{
        AccountEnabled    = [bool]$EnableUser
        DisplayName       = [string]$Row.DisplayName
        GivenName         = [string]$Row.GivenName
        Surname           = [string]$Row.Surname
        MailNickname      = $mailNickname
        UserPrincipalName = [string]$Row.UserPrincipalName
        UserType          = $userType
        UsageLocation     = $usageLocation

        PasswordProfile = @{
            Password                      = $plainTextPassword
            ForceChangePasswordNextSignIn = $forcePasswordChangeValue
        }
    }

    $optionalProperties = @(
        'City'
        'State'
        'StreetAddress'
        'PostalCode'
        'Country'
        'OfficeLocation'
        'Department'
        'JobTitle'
        'CompanyName'
    )

    foreach ($property in $optionalProperties) {
        if (-not [string]::IsNullOrWhiteSpace([string]$Row.$property)) {
            $request[$property] = [string]$Row.$property
        }
    }

    return [pscustomobject]@{
        Request              = $request
        PlainTextPassword    = $plainTextPassword
        PasswordWasGenerated = $passwordWasGenerated
    }
}

function New-EntraAccount {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(
            Mandatory,
            ValueFromPipeline
        )]
        [psobject]$Row
    )

    process {
        $userPrincipalName = [string]$Row.UserPrincipalName

        if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
            return [pscustomobject]@{
                UserPrincipalName = $null
                Id                = $null
                Status            = 'Failed: UserPrincipalName is missing.'
                ManagerStatus     = 'Not attempted'
                LicenceStatus     = 'Not attempted'
                TemporaryPassword = $null
            }
        }

        if (-not $PSCmdlet.ShouldProcess(
            $userPrincipalName,
            'Create Microsoft Entra ID user'
        )) {
            return
        }

        $user = $null
        $managerStatus = 'Not requested'
        $licenceStatus = 'Not requested'

        try {
            $convertedRequest = ConvertTo-UserRequest -Row $Row
            $request = $convertedRequest.Request

            Write-Host "Creating $userPrincipalName..." -ForegroundColor Cyan

            $user = New-MgUser @request

            $extensionAttributes = @{}

            if (-not [string]::IsNullOrWhiteSpace([string]$Row.ExtensionAttribute1)) {
                $extensionAttributes.ExtensionAttribute1 = [string]$Row.ExtensionAttribute1
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$Row.ExtensionAttribute2)) {
                $extensionAttributes.ExtensionAttribute2 = [string]$Row.ExtensionAttribute2
            }

            if ($extensionAttributes.Count -gt 0) {
                Update-MgUser `
                    -UserId $user.Id `
                    -OnPremisesExtensionAttributes $extensionAttributes `
                    -ErrorAction Stop
            }

            if (-not $user.Id) {
                throw "Microsoft Graph did not return an ID for '$userPrincipalName'."
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$Row.ManagerEmail)) {
                try {
                    $managerEmail = ([string]$Row.ManagerEmail).Trim()

                    $manager = Get-MgUser `
                        -UserId $managerEmail `
                        -Property Id, DisplayName, UserPrincipalName `
                        -ErrorAction Stop

                    if (-not $manager.Id) {
                        throw "Manager '$managerEmail' was not found."
                    }

                    $managerReference = @{
                        '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($manager.Id)"
                    }

                    Set-MgUserManagerByRef `
                        -UserId $user.Id `
                        -BodyParameter $managerReference `
                        -ErrorAction Stop

                    $managerStatus = "Assigned: $($manager.UserPrincipalName)"
                }
                catch {
                    $managerStatus = "Failed: $($_.Exception.Message)"
                    Write-Warning "User was created, but manager assignment failed for '$userPrincipalName': $($_.Exception.Message)"
                }
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$Row.LicenseSku)) {
                try {
                    $licenceSku = ([string]$Row.LicenseSku).Trim()

                    $sku = Get-MgSubscribedSku -All |
                        Where-Object {
                            $_.SkuPartNumber -eq $licenceSku
                        } |
                        Select-Object -First 1

                    if (-not $sku) {
                        throw "Licence SKU '$licenceSku' was not found in the tenant."
                    }

                    if ($null -ne $sku.PrepaidUnits.Enabled) {
                        $availableLicences =
                            [int]$sku.PrepaidUnits.Enabled -
                            [int]$sku.PrepaidUnits.Warning -
                            [int]$sku.ConsumedUnits

                        if ($availableLicences -le 0) {
                            throw "No licences are available for SKU '$licenceSku'."
                        }
                    }

                    $addLicences = @(
                        @{
                            SkuId = $sku.SkuId
                        }
                    )

                    Set-MgUserLicense `
                        -UserId $user.Id `
                        -AddLicenses $addLicences `
                        -RemoveLicenses @() `
                        -ErrorAction Stop |
                        Out-Null

                    $licenceStatus = "Assigned: $licenceSku"
                }
                catch {
                    $licenceStatus = "Failed: $($_.Exception.Message)"
                    Write-Warning "User was created, but licence assignment failed for '$userPrincipalName': $($_.Exception.Message)"
                }
            }

            return [pscustomobject]@{
                UserPrincipalName = $user.UserPrincipalName
                Id                = $user.Id
                Status            = 'Created'
                ManagerStatus     = $managerStatus
                LicenceStatus     = $licenceStatus
                TemporaryPassword = $convertedRequest.PlainTextPassword
            }
        }
        catch {
            return [pscustomobject]@{
                UserPrincipalName = $userPrincipalName
                Id                = if ($user) { $user.Id } else { $null }
                Status            = "Failed: $($_.Exception.Message)"
                ManagerStatus     = $managerStatus
                LicenceStatus     = $licenceStatus
                TemporaryPassword = $null
            }
        }
    }
}

function Import-EntraUserCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "CSV file not found: $Path"
    }

    $rows = @(Import-Csv -LiteralPath $Path)

    if ($rows.Count -eq 0) {
        throw "The CSV file contains no user records: $Path"
    }

    $requiredHeaders = @(
        'DisplayName'
        'GivenName'
        'Surname'
        'UserPrincipalName'
    )

    $csvHeaders = @($rows[0].PSObject.Properties.Name)

    foreach ($requiredHeader in $requiredHeaders) {
        if ($requiredHeader -notin $csvHeaders) {
            throw "The CSV file is missing the required '$requiredHeader' column."
        }
    }

    return $rows
}

function Show-Result {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [psobject[]]$Result
    )

    $Result |
        Format-Table `
            UserPrincipalName,
            Status,
            ManagerStatus,
            LicenceStatus,
            TemporaryPassword `
            -AutoSize
}

Connect-Entra

switch ($EntryPoint) {
    'Single' {
        $result = New-EntraAccount -Row (Get-UserInput)
        Show-Result -Result @($result)
    }

    'Csv' {
        if ([string]::IsNullOrWhiteSpace($CsvPath)) {
            $CsvPath = Read-Required -Prompt 'CSV path'
        }

        $rows = Import-EntraUserCsv -Path $CsvPath

        $results = foreach ($row in $rows) {
            New-EntraAccount -Row $row
        }

        Show-Result -Result @($results)
    }

    'Menu' {
        do {
            Write-Host ''
            Write-Host 'Microsoft Entra ID Account Creation' -ForegroundColor Cyan
            Write-Host '1. Create one account'
            Write-Host '2. Create accounts from CSV'
            Write-Host '3. Exit'

            $selection = Read-Host 'Select an option'

            switch ($selection) {
                '1' {
                    try {
                        $result = New-EntraAccount -Row (Get-UserInput)
                        Show-Result -Result @($result)
                    }
                    catch {
                        Write-Error $_.Exception.Message
                    }
                }

                '2' {
                    try {
                        $CsvPath = Read-Required -Prompt 'CSV path'
                        $rows = Import-EntraUserCsv -Path $CsvPath

                        $results = foreach ($row in $rows) {
                            New-EntraAccount -Row $row
                        }

                        Show-Result -Result @($results)
                    }
                    catch {
                        Write-Error $_.Exception.Message
                    }
                }

                '3' {
                    Write-Host 'Exiting.' -ForegroundColor Yellow
                    return
                }

                default {
                    Write-Warning 'Invalid option. Select 1, 2 or 3.'
                }
            }
        }
        while ($true)
    }
}