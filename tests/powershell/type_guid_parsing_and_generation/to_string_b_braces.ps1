# vybe-test: powershell/type_guid_parsing_and_generation/to_string_b_braces
$g = [guid]::Parse("d3b07384-d113-40a2-b94f-a496c152a3b1")
if ($g.ToString("B") -ne "{d3b07384-d113-40a2-b94f-a496c152a3b1}") {
    Write-Host "FAIL: ToString('B') failed"
    exit 1
}
Write-Host "PASS"
exit 0
