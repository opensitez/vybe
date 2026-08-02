<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_unhandled_match_error_php80
// origin: languages/php/tests/php/test_php_exception_types_spl_hierarchy.rs
// vybe-test-mode: compile

try {
    $x = 99;
    match ($x) {
        1 => "one",
        2 => "two",
    };
} catch (UnhandledMatchError $e) {
    echo "UnhandledMatchError: " . $e->getMessage();
}
