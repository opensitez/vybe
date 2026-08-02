<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_callable_type_hint_validation
// origin: languages/php/tests/php/test_php_callables_is_callable_call_user_func.rs
// vybe-test-mode: compile

function executeCallback(callable $cb, mixed ...$args) {
    return $cb(...$args);
}

echo executeCallback(fn($x, $y) => $x * $y, 6, 7);
