# vybe-test: powershell/string_padding_and_alignment/table_column_formatting_simulation
$col1 = "{0,-8}" -f "Name"
$col2 = "{0,6}" -f "Qty"
$row = "$col1|$col2"
if ($row -ne "Name    |   Qty") {
    Write-Host "FAIL: Table column alignment failed, got '$row'"
    exit 1
}
Write-Host "PASS"
exit 0
