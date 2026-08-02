<?php
// vybe-test: php/match_advanced/match_as_expression_in_call
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

function describe(string $s): string { return "[$s]"; }
$level = 3;
echo describe(match($level) {
    1 => 'low',
    2 => 'medium',
    3 => 'high',
    default => 'unknown',
});
