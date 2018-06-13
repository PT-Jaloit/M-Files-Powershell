[System.Reflection.Assembly]::LoadFrom("Interop.MFilesAPI.dll")

Function Get-MFilesVault {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [string]$Name,
        [parameter(Mandatory=$True)]
        [object[]]$Connection
    )

    $Connection.GetVaultConnection($Name)
}

Function Connect-MFilesVault {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$VaultConnection,
        [Parameter(Mandatory=$True)]
        [object]$AuthType,
        [Parameter(Mandatory=$False)]
        [string]$UserName,
        [Parameter(Mandatory=$False)]
        [string]$Password,
        [Parameter(Mandatory=$False)]
        [string]$Domain,
        [Parameter(Mandatory=$False)]
        [string]$SPN
    )
    
    If($AuthType -and !$UserName -and !$Password -and !$Domain -and !$SPN){
        $VaultConnection.LogInAsUser($AuthType)
    }
    If($AuthType -and $UserName -and !$Password -and !$Domain -and !$SPN){
        $VaultConnection.LogInAsUser($AuthType, $UserName)
    }
    If($AuthType -and $UserName -and $Password -and !$Domain -and !$SPN){
        $VaultConnection.LogInAsUser($AuthType, $UserName, $Password)
    }
    If($AuthType -and $UserName -and $Password -and $Domain -and !$SPN){
        $VaultConnection.LogInAsUser($AuthType, $UserName, $Password, $Domain)
    }
    If($AuthType -and $UserName -and $Password -and $Domain -and $SPN){
        $VaultConnection.LogInAsUser($AuthType, $UserName, $Password, $Domain, $SPN)
    }
}

Function Test-MFilesVaultConnection {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$VaultConnection
    )

    $Status = $VaultConnection.TestConnectionToVaultSilent()
    If($Status -eq 0){
        "Login succeeded"
    } ElseIf ($Status -eq 1){
        "Login failed"
    } Else {
        "Unknown error"
    }
}

Function Set-MFilesVaultConnection {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$VaultConnection
    )

    $VaultConnection.BindToVault(0, $true, $false)
}

Function Set-MFilesObjID {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [Long]$ObjID,
        [parameter(Mandatory=$True)]
        [Long]$ObjType
    )

    return $(Set-MFilesObjVer -ObjType $ObjType -ObjID $ObjID -Version -1).ObjID
}

Function Set-MFilesObjVer {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [Long]$ObjID,
        [parameter(Mandatory=$True)]
        [Long]$ObjType,
        [parameter(Mandatory=$True)]
        [Long]$Version
    )

    $ObjVer = New-Object -Com MFilesAPI.ObjVer
    $ObjVer.SetIDs($ObjType, $ObjID, $Version)
    return $ObjVer
}

Function Get-MFilesPropertyDefByID {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$Vault,
        [parameter(Mandatory=$True)]
        [Long]$Id
    )
    $Vault.PropertyDefOperations.GetPropertyDef($ID)
}


Function Get-MFilesObjectPropertiesOutput {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$Vault,
        [parameter(Mandatory=$True)]
        [object]$Properties
    )

    #$Return = @()
    $Return = New-Object –TypeName PSObject
    Foreach ($Property in $Properties) {
        #$Temp = New-Object –TypeName PSObject
        #$Temp | Add-Member -Name 'PropertyDef' -MemberType Noteproperty -Value $Property.PropertyDef
        #$Temp | Add-Member -Name 'PropertyName' -MemberType Noteproperty -Value $(Get-MFilesPropertyDefByID -Vault $Vault -Id $Property.PropertyDef).Name
        #$Temp | Add-Member -Name 'DisplayValue' -MemberType Noteproperty -Value $Property.Value.DisplayValue
        $Return | Add-Member -Name $(Get-MFilesPropertyDefByID -Vault $Vault -Id $Property.PropertyDef).Name -MemberType Noteproperty -Value $Property.Value.DisplayValue
        #$Return += $Temp
    }
    $Return
}

Function Get-MFilesObjectProperties {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$Vault,
        [parameter(Mandatory=$True)]
        [Long]$Id,
        [parameter(Mandatory=$True)]
        [Long]$Type
    )

    $ObjID = Set-MFilesObjID -ObjID $Id -ObjType $Type
    $Object = $Vault.ObjectOperations

    Get-MFilesObjectPropertiesOutput -Vault $Vault -Properties $Object.GetLatestObjectVersionAndProperties($ObjID, $true).Properties
}

Function Get-MFilesSearchCondition {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [String]$Expression,
        [parameter(Mandatory=$True)]
        [object]$ConditionType,
        [parameter(Mandatory=$True)]
        [String]$Value,
        [parameter(Mandatory=$True)]
        [object]$DataType,
        [parameter(Mandatory=$False)]
        [object]$Property,
        [parameter(Mandatory=$False)]
        [object]$Status
    )
    
    $Condition = New-Object -Com MFilesAPI.SearchCondition
    $Condition.ConditionType = $ConditionType
    $Condition.TypedValue.SetValue($DataType, $Value)

    switch ( $Expression )
    {
        "Status"   { $Condition.Expression.SetStatusValueExpression($Status) }
        "Property" { $Condition.Expression.DataPropertyValuePropertyDef = $Property }
    }
    $Condition
}


Function Get-MFilesSearch {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$Vault,
        [parameter(Mandatory=$True)]
        [object]$Conditions
    )
    
    $Vault.ObjectSearchOperations.SearchForObjectsByConditions($Conditions, $([MFilesAPI.MFSearchFlags]::MFSearchFlagNone), $false)
}

Function Get-MFilesSearchMore {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$Vault,
        [parameter(Mandatory=$True)]
        [object]$Conditions
    )
    
    $Vault.ObjectSearchOperations.SearchForObjectsByConditions($Conditions, $([MFilesAPI.MFSearchFlags]::MFSearchFlagNone), $false).MoreResults
}

Function Get-MFilesSearchCount {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$True)]
        [object[]]$Vault,
        [parameter(Mandatory=$True)]
        [object]$Conditions
    )
    
    $Vault.ObjectSearchOperations.GetObjectCountInSearch($Conditions, $([MFilesAPI.MFSearchFlags]::MFSearchFlagNone))
}