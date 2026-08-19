# vybe-test: powershell/whitespace_and_line_rules/nonbreaking_space_rejected_identifier
$nb = [char]0x00A0
$code = "va${nb}r = 1"
$errors = @()
$tokens = [System.Management.Automation.PSParser]::Tokenize($code, [ref]$errors)
$tokensText = ($tokens | ForEach-Object Content)
if ($errors.Count -ne 0) {
    Write-Host 'FAIL: non-breaking space produced parser errors'
    exit 1
}

if (($tokensText.Count -lt 3) -or ($tokensText[0] -ne 'va') -or ($tokensText[1] -ne 'r')) {
    Write-Host "FAIL: non-breaking space did not split identifier as expected: $($tokensText -join ',')"
    exit 1
}

Write-Host 'PASS'
exit 0
