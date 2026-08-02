<?php
// vybe-test: php/functional_patterns/memoize_expensive_call
// origin: languages/php/tests/php/test_functional_patterns.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

function memoize(callable $fn): Closure {
    $cache = [];
    return function() use ($fn, &$cache) {
        $key = serialize(func_get_args());
        if (!array_key_exists($key, $cache)) {
            $cache[$key] = $fn(...func_get_args());
        }
        return $cache[$key];
    };
}
$calls = 0;
$expensiveFib = memoize(function(int $n) use (&$calls): int {
    $calls++;
    if ($n <= 1) return $n;
    $fib = 0; $a = 0; $b = 1;
    for ($i = 2; $i <= $n; $i++) { $fib = $a + $b; $a = $b; $b = $fib; }
    return $fib;
});
echo $expensiveFib(10) . ',' . $expensiveFib(10) . ',calls:' . $calls;

__vybe_check(ob_get_clean(), "55,55,calls:1");
