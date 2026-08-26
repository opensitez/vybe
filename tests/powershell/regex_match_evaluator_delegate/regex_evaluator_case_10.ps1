# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_10
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_10", "word_10", $eval)
if ($res -ne "[word_10]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
