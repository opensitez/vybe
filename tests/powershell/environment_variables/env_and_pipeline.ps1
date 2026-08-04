# vybe-test: powershell/environment_variables/env_and_pipeline
$env:PIPE_VAR = 'p'
if ((1 | ForEach-Object { $env:PIPE_VAR }) -eq 'p') { exit 0 }
exit 1
