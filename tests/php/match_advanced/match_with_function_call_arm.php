<?php
// vybe-test: php/match_advanced/match_with_function_call_arm
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

function heavy(): string { return 'computed'; }
$flag = true;
$result = match($flag) {
    true  => heavy(),
    false => 'skipped',
};
echo $result;
