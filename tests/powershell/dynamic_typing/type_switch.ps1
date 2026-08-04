# vybe-test: powershell/dynamic_typing/type_switch
$value = '5'
if (($value.GetType().Name) -eq 'String') { exit 0 }
exit 1
