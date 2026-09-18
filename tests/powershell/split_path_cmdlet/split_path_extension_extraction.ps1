# vybe-test: powershell/split_path_cmdlet/split_path_extension_extraction
# Split-Path -Extension extracts the file extension including the leading dot
$res = Split-Path "/dir/document.docx" -Extension

if ($res -ne ".docx") {
    Write-Host "FAIL: expected '.docx', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
