# vybe-test: powershell/string_splitting/split_limit
if (("a,b,c" -split ',', 2).Count -eq 2) { exit 0 }
exit 1
