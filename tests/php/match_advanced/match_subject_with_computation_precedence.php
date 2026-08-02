<?php
// vybe-test: php/match_advanced/match_subject_with_computation_precedence
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$n = 3;
$v = match (1 + 2 * $n) {
    5 => 'ok',
    7 => 'oops',
    default => 'other',
};
echo $v;
