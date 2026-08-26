# vybe-test: powershell/regex_match_evaluator_delegate/regex_replace_with_match_evaluator_delegate
$text = "apple 10 banana 20 cherry 30"
$evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    return ([int]$m.Value * 2).ToString()
}
$res = [regex]::Replace($text, "\d+", $evaluator)
if ($res -ne "apple 20 banana 40 cherry 60") { Write-Host "FAIL: MatchEvaluator replace failed, got '$res'"; exit 1 }
Write-Host "PASS"; exit 0
