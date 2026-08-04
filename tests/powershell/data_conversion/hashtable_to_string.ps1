# vybe-test: powershell/data_conversion/hashtable_to_string
$x = @{ a = 1 }
if ($x.ToString().Contains('System.Collections.Hashtable')) { exit 0 }
exit 1
