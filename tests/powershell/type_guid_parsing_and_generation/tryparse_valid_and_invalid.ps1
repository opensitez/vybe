# vybe-test: powershell/type_guid_parsing_and_generation/tryparse_valid_and_invalid
$g = [guid]::Empty
$v = [guid]::TryParse("d3b07384-d113-40a2-b94f-a496c152a3b1", [ref]$g)
$inv = [guid]::TryParse("not-a-guid", [ref]$g)
if (-not $v -or $inv) {
    Write-Host "FAIL: TryParse check failed"
    exit 1
}
Write-Host "PASS"
exit 0
