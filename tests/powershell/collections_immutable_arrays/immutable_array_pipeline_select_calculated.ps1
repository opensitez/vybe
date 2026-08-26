# vybe-test: powershell/collections_immutable_arrays/immutable_array_pipeline_select_calculated
$arr = [System.Collections.Immutable.ImmutableArray]::Create([string[]]@("hello", "world"))
$res = @($arr | Select-Object @{ N = "Upper"; E = { $_.ToUpper() } })
if ($res[0].Upper -ne "HELLO") { Write-Host "FAIL: Pipeline select failed"; exit 1 }
Write-Host "PASS"; exit 0
