function Get-UpperName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        throw "Name is required"
    }
    return $Name.ToUpper()
}

Write-Output (Get-UpperName -Name "powershell")
Write-Output (Get-UpperName -Name "compiler")
