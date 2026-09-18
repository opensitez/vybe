# vybe-test: powershell/extended_type_system/ets_default_display_property_single
# -DefaultDisplayProperty designates the single default column for formatters (e.g. Format-Wide)
class SingleDisplayCandidate {
    [string]$Title = "Main Title"
    [string]$Description = "Long description body"
}

Update-TypeData -TypeName SingleDisplayCandidate -DefaultDisplayProperty Title -Force
$td = Get-TypeData -TypeName SingleDisplayCandidate

if ($td.DefaultDisplayProperty -ne "Title") {
    Write-Host "FAIL: expected DefaultDisplayProperty 'Title', got: '$($td.DefaultDisplayProperty)'"
    exit 1
}

Write-Host "PASS"
exit 0
