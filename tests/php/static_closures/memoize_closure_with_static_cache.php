<?php
// vybe-test: php/static_closures/memoize_closure_with_static_cache
// origin: languages/php/tests/php/test_static_closures.rs

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
        $args = func_get_args();
        $key = serialize($args);
        if (!isset($cache[$key])) {
            $cache[$key] = $fn(...$args);
        }
        return $cache[$key];
    };
}
$expensive = memoize(function(int $n): int { return $n * $n; });
echo $expensive(4);
echo $expensive(4);

__vybe_check(ob_get_clean(), "1616");
