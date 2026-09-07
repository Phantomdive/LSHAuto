Import-Module ActiveDirectory

$csv = Import-Csv .\extAttrs.csv

foreach ($row in $csv) {
    Set-ADUser `
        -Identity $row.UPN `
        -Replace @{ extensionAttribute2 = $row.extensionAttribute2 }
}