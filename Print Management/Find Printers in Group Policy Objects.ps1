# A Powershell script to find the Group Policy Objects that printers have been deployed to
# Exports the results to a CSV file at $exportPath

$exportPath = "C:\Temp\gpo_printers.csv"

$printers = Get-ADObject -LDAPFilter "(objectClass=msPrint-ConnectionPolicy)" -SearchBase "CN=Policies,CN=System,DC=andshrew,DC=com" -Properties *

foreach ($printer in $printers) {
    # Get the GUID of the Group Policy Object, then find it's name
    $guid = $printer.DistinguishedName.Substring($printer.DistinguishedName.LastIndexOf("{"),38)
    $gpo = Get-GPO -Guid $guid
    Add-Member -InputObject $printer -MemberType NoteProperty -Name "GPO" -Value $gpo.DisplayName -Force
}

$printers | select printerName, uNCName, whenChanged, whenCreated, GPO | Export-Csv -NoTypeInformation -Path $exportPath