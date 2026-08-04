# vybe-test: powershell/environment_variables/env_in_script
$env:TEST_SCRIPT = 's'
function Test { $env:TEST_SCRIPT = 't' }
Test
if ($env:TEST_SCRIPT -eq 't') { exit 0 }
exit 1
