<?php
// vybe-test: php/functional_style/memoization_static_cache
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function memoize(callable $fn): callable {
    $cache = [];
    return function() use ($fn, &$cache) {
        $key = serialize(func_get_args());
        if (!isset($cache[$key])) {
            $cache[$key] = $fn(...func_get_args());
        }
        return $cache[$key];
    };
}
$expensiveAdd = memoize(fn($a, $b) => $a + $b);
echo $expensiveAdd(3, 4);
echo $expensiveAdd(3, 4);
