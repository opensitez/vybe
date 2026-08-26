# vybe-test: powershell/regex_match_evaluator_delegate/regex_match_evaluator_uppercase_transform
$text = "hello world from powershell"
$evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    return $m.Value.ToUpper()
}
$res = [regex]::Replace($text, "\w+", $evaluator)
if ($res -ne "HELLO WORLD FROM POWERSHELL") { Write-Host "FAIL: MatchEvaluator uppercase failed, got '$res'"; exit 1 }
Write-Host "PASS"; exit 0
