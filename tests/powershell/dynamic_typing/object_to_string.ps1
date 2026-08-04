# vybe-test: powershell/dynamic_typing/object_to_string
$x = [pscustomobject]@{ Name = 'v' }
$x = $x.Name
if ($x -eq 'v') { exit 0 }
exit 1
