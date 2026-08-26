# vybe-test: powershell/csv_delimiter_and_quoting/custom_tab_delimiter_import
$csv = "Name`tScore`nAlpha`t100`nBeta`t95"
$rows = @($csv | ConvertFrom-Csv -Delimiter "`t")
if ($rows.Length -ne 2 -or $rows[0].Name -ne "Alpha" -or $rows[0].Score -ne "100") {
    Write-Host "FAIL: Tab delimiter import failed"
    exit 1
}
Write-Host "PASS"
exit 0
