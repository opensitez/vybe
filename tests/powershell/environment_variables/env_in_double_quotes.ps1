# vybe-test: powershell/environment_variables/env_in_double_quotes
$env:DOUBLE = 'x'
if ("$env:DOUBLE" -eq 'x') { exit 0 }
exit 1
