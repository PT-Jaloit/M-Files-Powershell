# M-Files-Powershell

## Connecting to Document Vault 
```powershell
# Using M-Files Client
$MFClient = New-Object -COM MFilesAPI.MFilesClientApplication

# Select Document Vault
$VaultConnection = Get-MFilesVault -Name "VAULT" -Connection $MFClient

# Test Connection to Document Vault
Test-MFilesVaultConnection -VaultConnection $VaultConnection

# Connect to Document Vault using different credentials
$Vault = Connect-MFilesVault -VaultConnection $VaultConnection -AuthType $([MFilesAPI.MFAuthType]::MFAuthTypeSpecificWindowsUser) -UserName "username" -Password "password" -Domain "domain"

# Connect to Document Vault
$Vault = Set-MFilesVaultConnection -VaultConnection $VaultConnection
```

## Get Object Properties by ID and object type
```powershell
Get-MFilesObjectProperties -Vault $Vault -Id 1 -Type 0
```

## Use M-Files Search with conditions
```powershell
# Multiple search conditions
$Conditions = New-Object -Com MFilesAPI.SearchConditions
# Document name or title Contains XXX
$Conditions.Add(1, $(Get-MFilesSearchCondition -Expression "Property" -ConditionType $([MFilesAPI.MFConditionType]::MFConditionTypeContains)  -Value "Test" -DataType $([MFilesAPI.MFDataType]::MFDatatypeText) -Property "0"))
# Class equals Document
$Conditions.Add(2, $(Get-MFilesSearchCondition -Expression "Property" -ConditionType $([MFilesAPI.MFConditionType]::MFConditionTypeEqual)  -Value "0" -DataType $([MFilesAPI.MFDataType]::MFDatatypeLookup) -Property "100"))

# Get search results
$Results = Get-MFilesSearch -Vault $Vault -Conditions $Conditions

# Return name or title and last modified date.
Foreach ($Item in $Results){
    Get-MFilesObjectProperties -Vault $Vault -Id $Item.ObjVer.ID -Type $Item.ObjVer.Type | FT "Nimi tai otsikko", "Viimeksi muokattu"
}
```