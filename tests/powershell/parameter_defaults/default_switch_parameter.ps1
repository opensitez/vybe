# vybe-test: powershell/parameter_defaults/default_switch_parameter
function Test { param([switch]$Flag = $false) if ($Flag) { 'yes' } else { 'no' } }
if ((Test) -eq 'no') { exit 0 }
exit 1
