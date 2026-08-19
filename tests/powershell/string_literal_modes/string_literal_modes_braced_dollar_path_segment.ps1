# vybe-test: powershell/string_literal_modes/braced_dollar_path_segment
$drive = 'C'
$folder = 'temp'
$path = "${drive}:\${folder}"
if ($path -ne 'C:\temp') {
    Write-Host "FAIL: expected C:\\temp, got '$path'"
    exit 1
}

Write-Host 'PASS'
exit 0
