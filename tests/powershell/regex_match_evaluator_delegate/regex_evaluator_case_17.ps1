# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_17
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_17", "word_17", $eval)
if ($res -ne "[word_17]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
