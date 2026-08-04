# vybe-test: powershell/environment_variables/env_persistence
$env:TEST_PERSIST = 'persist'
if ($env:TEST_PERSIST -eq 'persist') { exit 0 }
exit 1
