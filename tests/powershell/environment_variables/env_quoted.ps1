# vybe-test: powershell/environment_variables/env_quoted
$env:QUOTE = 'a b'
if ($env:QUOTE -eq 'a b') { exit 0 }
exit 1
