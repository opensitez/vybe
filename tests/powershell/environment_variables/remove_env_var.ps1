# vybe-test: powershell/environment_variables/remove_env_var
$env:TEST_REMOVE = 'x'
Remove-Item Env:TEST_REMOVE -ErrorAction SilentlyContinue
if ($env:TEST_REMOVE -eq $null) { exit 0 }
exit 1
