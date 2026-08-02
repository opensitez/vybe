<?php
// vybe-test: php/closures_advanced/closure_returned_from_function
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

function multiplier(int $factor): Closure {
    return fn(int $n) => $n * $factor;
}
$double = multiplier(2);
$triple = multiplier(3);
echo $double(5) . ',' . $triple(5);
