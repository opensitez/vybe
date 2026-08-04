# vybe-test: powershell/string_splitting/split_empty
if (("" -split ',').Count -eq 1) { exit 0 }
exit 1
