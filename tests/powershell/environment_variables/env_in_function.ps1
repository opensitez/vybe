# vybe-test: powershell/environment_variables/env_in_function
function Test { $env:TEST_FN = 'f' }
Test
if ($env:TEST_FN -eq 'f') { exit 0 }
exit 1
