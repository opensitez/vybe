<?php
// vybe-test: php/closures_advanced/closure_from_callable_compose
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

function compose(callable ...$fns): Closure {
    return function($v) use ($fns) {
        return array_reduce(
            array_reverse($fns),
            fn($carry, $fn) => $fn($carry),
            $v
        );
    };
}
$process = compose(
    Closure::fromCallable('strtoupper'),
    Closure::fromCallable('trim')
);
echo $process("  hello world  ");
