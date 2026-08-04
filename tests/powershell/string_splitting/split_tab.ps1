# vybe-test: powershell/string_splitting/split_tab
if (("a	b" -split '\t').Count -eq 2) { exit 0 }
exit 1
