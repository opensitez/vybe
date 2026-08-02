<?php
// vybe-test: php/phase2/match_with_calls
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$action = 'greet';
match($action) {
    'greet' => 'hello',
    'bye' => 'goodbye',
    default => 'unknown'
};
