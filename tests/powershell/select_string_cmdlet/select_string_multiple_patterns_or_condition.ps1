# vybe-test: powershell/select_string_cmdlet/select_string_multiple_patterns_or_condition
# Supplying an array of patterns evaluates an OR condition across the patterns
$animals = @("cat", "parrot", "dog", "hamster")
$matches = @($animals | Select-String -Pattern @("cat", "dog"))

if ($matches.Count -ne 2) {
    Write-Host "FAIL: expected 2 matches for multiple patterns, got $($matches.Count)"
    exit 1
}

$matchedLines = @($matches | ForEach-Object { $_.Line })
if ($matchedLines -notcontains "cat" -or $matchedLines -notcontains "dog") {
    Write-Host "FAIL: expected 'cat' and 'dog', got: @($($matchedLines -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
