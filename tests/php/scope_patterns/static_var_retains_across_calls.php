<?php
// vybe-test: php/scope_patterns/static_var_retains_across_calls
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

function increment(): int {
    static $n = 0;
    return ++$n;
}
echo increment();
echo increment();
echo increment();
