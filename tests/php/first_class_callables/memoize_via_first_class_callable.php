<?php
// vybe-test: php/first_class_callables/memoize_via_first_class_callable
// origin: languages/php/tests/php/test_first_class_callables.rs

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

function memoize(callable $fn): callable {
    $cache = [];
    return function() use ($fn, &$cache) {
        $key = serialize(func_get_args());
        if (!array_key_exists($key, $cache)) {
            $cache[$key] = $fn(...func_get_args());
        }
        return $cache[$key];
    };
}
$callCount = 0;
function expensiveCompute(int $n) use (&$callCount): int {
    $callCount++;
    return $n * $n;
}
$memoized = memoize(expensiveCompute(...));
echo $memoized(4) . "\n";
echo $memoized(4) . "\n";
echo $memoized(5) . "\n";
echo $callCount . "\n";

__vybe_check(ob_get_clean(), "16\n16\n25\n2");
