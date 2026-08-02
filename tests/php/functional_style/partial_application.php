<?php
// vybe-test: php/functional_style/partial_application
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function multiply(int $a, int $b): int { return $a * $b; }
function partial(callable $fn, ...$partialArgs): callable {
    return fn(...$rest) => $fn(...$partialArgs, ...$rest);
}
$double = partial('multiply', 2);
$triple = partial('multiply', 3);
echo $double(5);
echo $triple(5);
