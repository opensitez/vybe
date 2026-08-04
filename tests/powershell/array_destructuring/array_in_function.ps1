# vybe-test: powershell/array_destructuring/array_in_function
function Test { param($a,$b) return "$a,$b" }
if ((Test 1 2) -eq '1,2') { exit 0 }
exit 1
