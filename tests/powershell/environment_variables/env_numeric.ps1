# vybe-test: powershell/environment_variables/env_numeric
$env:NUM = 5
if ($env:NUM -eq '5') { exit 0 }
exit 1
