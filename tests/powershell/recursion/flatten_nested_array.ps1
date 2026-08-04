# vybe-test: powershell/recursion/flatten_nested_array
function Flatten([object[]]$arr) {
    $result = [System.Collections.Generic.List[int]]::new()
    foreach ($item in $arr) {
        if ($item -is [array]) {
            foreach ($sub in (Flatten $item)) { $result.Add($sub) }
        } else {
            $result.Add($item)
        }
    }
    return $result.ToArray()
}
$nested = @(1, @(2, 3), @(4, @(5, 6)))
$flat = Flatten $nested
$expected = @(1,2,3,4,5,6)
for ($i = 0; $i -lt 6; $i++) {
    if ($flat[$i] -ne $expected[$i]) {
        Write-Host "FAIL: flat[$i] = $($flat[$i]), expected $($expected[$i])"
        exit 1
    }
}
Write-Host "PASS"
exit 0
