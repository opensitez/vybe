# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_9
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_9", "word_9", $eval)
if ($res -ne "[word_9]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
