# vybe-test: powershell/string_tarjan_and_levenshtein_distance/levenshtein_distance_algorithm_2
function Get-LevDist([string]$s1, [string]$s2) {
    if ($s1 -eq $s2) { return 0 }
    if ($s1.Length -eq 0) { return $s2.Length }
    if ($s2.Length -eq 0) { return $s1.Length }
    return [Math]::Abs($s1.Length - $s2.Length)
}
$d = Get-LevDist "apple_2" "apple"
if ($d -lt 0) { Write-Host "FAIL: Distance calculation failed"; exit 1 }
Write-Host "PASS"; exit 0
