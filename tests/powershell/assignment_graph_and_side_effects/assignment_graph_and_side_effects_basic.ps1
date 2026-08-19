# vybe-test: powershell/assignment_graph_and_side_effects/basic
$state = [ordered]@{ a = 1; b = 2 }
$state.a += $state.b
$state.b = $state.a * 2
$state.c = $state.a + $state.b

if ($state.a -ne 3 -or $state.b -ne 6 -or $state.c -ne 9) {
    Write-Host "FAIL: assignment graph mismatch a=$($state.a) b=$($state.b) c=$($state.c)"
    exit 1
}

Write-Host 'PASS'
exit 0
