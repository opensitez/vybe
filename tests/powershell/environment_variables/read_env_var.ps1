# vybe-test: powershell/environment_variables/read_env_var
$env:TEST_VAR = 'yes'
if ($env:TEST_VAR -eq 'yes') { exit 0 }
exit 1
