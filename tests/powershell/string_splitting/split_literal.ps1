# vybe-test: powershell/string_splitting/split_literal
if (("a.b.c" -split '\.').Count -eq 3) { exit 0 }
exit 1
