# vybe-test: powershell/string_splitting/split_string
if (("a--b" -split '--').Count -eq 2) { exit 0 }
exit 1
