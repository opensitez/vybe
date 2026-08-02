<?php
// vybe-test: php/closures_advanced/partial_application
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

function partial(callable $fn, mixed ...$partialArgs): Closure {
    return function() use ($fn, $partialArgs) {
        $args = array_merge($partialArgs, func_get_args());
        return $fn(...$args);
    };
}
function add(int $a, int $b, int $c): int { return $a + $b + $c; }
$add10 = partial('add', 10);
$add10and20 = partial('add', 10, 20);
echo $add10(5, 3);
echo $add10and20(7);
