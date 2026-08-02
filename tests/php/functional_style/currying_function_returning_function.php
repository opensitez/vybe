<?php
// vybe-test: php/functional_style/currying_function_returning_function
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function curry(callable $fn): callable {
    $arity = (new ReflectionFunction(Closure::fromCallable($fn)))->getNumberOfParameters();
    $accumulate = function(array $args) use ($fn, $arity, &$accumulate): mixed {
        if (count($args) >= $arity) {
            return $fn(...$args);
        }
        return fn(...$more) => $accumulate(array_merge($args, $more));
    };
    return fn(...$args) => $accumulate($args);
}
$add = curry(fn($a, $b) => $a + $b);
$add5 = $add(5);
echo $add5(3);
echo $add(10)(20);
