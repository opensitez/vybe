# vybe-test: powershell/datetime/datetime_format_string
$d = [DateTime]::new(2024, 7, 4)
$fmt = $d.ToString("yyyy-MM-dd")
if ($fmt -ne "2024-07-04") {
    Write-Host "FAIL: expected '2024-07-04', got '$fmt'"
    exit 1
}
$short = $d.ToString("MMM dd")
if ($short -ne "Jul 04") {
    Write-Host "FAIL: expected 'Jul 04', got '$short'"
    exit 1
}
Write-Host "PASS"
exit 0
