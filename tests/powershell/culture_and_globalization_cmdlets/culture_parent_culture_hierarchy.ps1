# vybe-test: powershell/culture_and_globalization_cmdlets/culture_parent_culture_hierarchy
# CultureInfo exposes a Parent property linking specific cultures back up to neutral parent cultures
$fr = Get-Culture -Name "fr-FR"

if ($fr.Parent.Name -ne "fr") {
    Write-Host "FAIL: expected parent culture 'fr', got: '$($fr.Parent.Name)'"
    exit 1
}

# Neutral culture's parent is invariant (empty Name)
if ($fr.Parent.Parent.Name -ne "") {
    Write-Host "FAIL: expected invariant grandparent (empty name), got: '$($fr.Parent.Parent.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
