<?php
// vybe-test: php/advanced_closures/closure_factory_returns_closure
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$makeAdder = function(int $base): Closure {
    return function(int $x) use ($base): int { return $base + $x; };
};
$add100 = $makeAdder(100);
echo $add100(42);
