# vybe-test: powershell/string_splitting/split_missing
if (("a" -split ',').Count -eq 1) { exit 0 }
exit 1
