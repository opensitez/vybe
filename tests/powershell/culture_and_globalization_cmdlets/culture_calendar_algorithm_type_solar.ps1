# vybe-test: powershell/culture_and_globalization_cmdlets/culture_calendar_algorithm_type_solar
# The Calendar property of the current culture reflects its algorithmic basis (Gregorian = SolarCalendar)
$current = Get-Culture
$algoType = $current.Calendar.AlgorithmType.ToString()

if ($algoType -ne "SolarCalendar") {
    Write-Host "FAIL: expected SolarCalendar, got: '$algoType'"
    exit 1
}

Write-Host "PASS"
exit 0
