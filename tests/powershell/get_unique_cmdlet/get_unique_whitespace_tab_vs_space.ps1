# vybe-test: powershell/get_unique_cmdlet/get_unique_whitespace_tab_vs_space
# Tab character `t and space character ' ' are treated as distinct strings by Get-Unique -AsString
$items = @("`t", " ", " ")
$unique = @($items | Get-Unique -AsString)

if ($unique.Count -ne 2) {
    Write-Host "FAIL: expected 2 elements (tab and space), got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne "`t" -or $unique[1] -ne " ") {
    Write-Host "FAIL: whitespace character distinction failed"
    exit 1
}

Write-Host "PASS"
exit 0
