<?php
// vybe-test: php/exception_types/type_error_thrown
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

function strictAdd(int $a, int $b): int { return $a + $b; }
try {
    throw new TypeError('argument must be int');
} catch (TypeError $e) {
    echo $e->getMessage();
}
