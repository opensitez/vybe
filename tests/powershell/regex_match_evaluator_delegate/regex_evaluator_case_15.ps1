# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_15
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_15", "word_15", $eval)
if ($res -ne "[word_15]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
