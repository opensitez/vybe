# vybe-test: powershell/string_splitting/split_regex
if (("a1b2c" -split '\d').Count -eq 3) { exit 0 }
exit 1
