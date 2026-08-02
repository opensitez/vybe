<?php
// vybe-test: php/closures_advanced/currying_with_closures
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

function curry(callable $fn): Closure {
    $arity = (new ReflectionFunction(Closure::fromCallable($fn)))->getNumberOfParameters();
    $accumulate = function(array $args) use ($fn, $arity, &$accumulate): mixed {
        if (count($args) >= $arity) return $fn(...$args);
        return function() use ($args, $accumulate) {
            return $accumulate(array_merge($args, func_get_args()));
        };
    };
    return function() use ($accumulate) { return $accumulate(func_get_args()); };
}
$add = curry(fn(int $a, int $b, int $c) => $a + $b + $c);
echo $add(1)(2)(3);
