# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_5
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_5", "word_5", $eval)
if ($res -ne "[word_5]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
