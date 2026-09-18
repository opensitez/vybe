# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_nested_objects
# Nested hierarchical PSCustomObject graphs are preserved through CliXml
$tmp = [System.IO.Path]::GetTempFileName()
$tree = [pscustomobject]@{
    Department = "Engineering"
    Manager = [pscustomobject]@{
        Name = "Alice"
        Level = 5
    }
}

$tree | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored.Department -ne "Engineering") {
    Write-Host "FAIL: top-level Department mismatch, got: '$($restored.Department)'"
    exit 1
}

if ($restored.Manager.Name -ne "Alice" -or $restored.Manager.Level -ne 5) {
    Write-Host "FAIL: nested Manager mismatch, got: $($restored.Manager.Name), $($restored.Manager.Level)"
    exit 1
}

Write-Host "PASS"
exit 0
