# vybe-test: powershell/environment_variables/set_env_var
$env:TEST_SET = 'ok'
if ($env:TEST_SET -eq 'ok') { exit 0 }
exit 1
