# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_custom_object
# Export-Clixml serializes PSCustomObject instances and Import-Clixml restores their NoteProperties
$tmp = [System.IO.Path]::GetTempFileName()
$orig = [pscustomobject]@{ Title = "PowerShell CookBook"; Cost = 45 }

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored.Title -ne "PowerShell CookBook") {
    Write-Host "FAIL: Title property mismatch, got: '$($restored.Title)'"
    exit 1
}

if ($restored.Cost -ne 45) {
    Write-Host "FAIL: Cost property mismatch, got: $($restored.Cost)"
    exit 1
}

Write-Host "PASS"
exit 0
