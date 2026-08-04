# vybe-test: powershell/string_splitting/split_string_method
if (("a,b".Split(',')).Count -eq 2) { exit 0 }
exit 1
