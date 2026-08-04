# vybe-test: powershell/collection_initializers/nested_hash
$x = @{ outer = @{ inner = 'ok' } }
if ($x.outer.inner -eq 'ok') { exit 0 }
exit 1
