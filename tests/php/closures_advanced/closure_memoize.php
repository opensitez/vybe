<?php
// vybe-test: php/closures_advanced/closure_memoize
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

function memoize(Closure $fn): Closure {
    $cache = [];
    return function() use ($fn, &$cache) {
        $key = serialize(func_get_args());
        if (!array_key_exists($key, $cache)) {
            $cache[$key] = $fn(...func_get_args());
        }
        return $cache[$key];
    };
}
$fib = memoize(function(int $n) use (&$fib): int {
    if ($n <= 1) return $n;
    return $fib($n - 1) + $fib($n - 2);
});
echo $fib(10);
