<?php
// vybe-test: php/scope_patterns/static_var_in_recursion
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

function depth(): int {
    static $calls = 0;
    $calls++;
    if ($calls < 4) depth();
    return $calls;
}
echo depth();
