# vybe-test: powershell/format_table_engine/table_custom_psnote_properties
$obj = [pscustomobject]@{ StaticBase = "BaseVal" }

# Attaching dynamic NoteProperty
$obj | Add-Member -NotePropertyName "DynamicScore" -NotePropertyValue 777

$output = $obj | Format-Table -Property StaticBase, DynamicScore | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The dynamically added note property must be formatted as a valid table column
if (-not ($output -match "DynamicScore" -and $output -match "777")) {
    Write-Host "FAIL: dynamic NoteProperty failed to format in table: $output"
    exit 1
}

Write-Host "PASS"
exit 0
