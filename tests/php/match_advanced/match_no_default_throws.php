<?php
// vybe-test: php/match_advanced/match_no_default_throws
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

try {
    $x = 42;
    $r = match($x) {
        1 => 'one',
        2 => 'two',
    };
} catch (\UnhandledMatchError $e) {
    echo 'unhandled match';
}
