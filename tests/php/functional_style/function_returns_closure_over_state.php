<?php
// vybe-test: php/functional_style/function_returns_closure_over_state
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function makeAccumulator(int $initial = 0): callable {
    $total = $initial;
    return function(int $n) use (&$total): int {
        $total += $n;
        return $total;
    };
}
$acc = makeAccumulator(10);
echo $acc(5);
echo $acc(3);
echo $acc(2);
