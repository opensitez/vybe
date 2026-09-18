# vybe-test: powershell/extended_type_system/ets_default_display_property_set
# -DefaultDisplayPropertySet defines the default visible columns in formatters for a type
class DisplayCandidate {
    [string]$PublicColA = "A"
    [string]$PublicColB = "B"
    [string]$InternalSecret = "Hidden"
}

Update-TypeData -TypeName DisplayCandidate -DefaultDisplayPropertySet PublicColA, PublicColB -Force
$td = Get-TypeData -TypeName DisplayCandidate

if ($null -eq $td.DefaultDisplayPropertySet) {
    Write-Host "FAIL: DefaultDisplayPropertySet was not registered"
    exit 1
}

$props = $td.DefaultDisplayPropertySet.ReferencedProperties
if ($props -notcontains "PublicColA" -or $props -notcontains "PublicColB") {
    Write-Host "FAIL: referenced properties mismatch: @($($props -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
