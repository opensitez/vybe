<?php
// vybe-test: php/exception_types/invalid_argument_exception_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

function divide(int $a, int $b): float {
    if ($b === 0) throw new InvalidArgumentException('divisor cannot be zero');
    return $a / $b;
}
try { divide(10, 0); } catch (InvalidArgumentException $e) { echo $e->getMessage(); }
