# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_12
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_12", "word_12", $eval)
if ($res -ne "[word_12]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
