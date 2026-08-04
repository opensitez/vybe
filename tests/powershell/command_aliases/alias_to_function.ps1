# vybe-test: powershell/command_aliases/alias_to_function
function Test-Me { param($x) return $x }
Set-Alias tm Test-Me
if ((tm 'ok') -eq 'ok') { exit 0 }
exit 1
