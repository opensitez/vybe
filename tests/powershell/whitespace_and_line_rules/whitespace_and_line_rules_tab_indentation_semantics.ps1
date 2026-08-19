# vybe-test: powershell/whitespace_and_line_rules/tab_indentation_semantics
if ($true) {
	$value = 5
	if ($value -eq 5) {
		$value = $value + 1
	}
}

if ($value -ne 6) {
    Write-Host "FAIL: tab indentation changed block execution: $value"
    exit 1
}

Write-Host 'PASS'
exit 0
