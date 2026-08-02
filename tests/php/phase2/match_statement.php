<?php
// vybe-test: php/phase2/match_statement
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$x = 2;
match($x) {
    1 => 'one',
    2 => 'two',
    default => 'other'
};
