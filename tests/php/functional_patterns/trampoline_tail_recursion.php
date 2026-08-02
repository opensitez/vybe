<?php
// vybe-test: php/functional_patterns/trampoline_tail_recursion
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

function trampoline(callable $fn): Closure {
    return function() use ($fn) {
        $result = $fn(...func_get_args());
        while (is_callable($result)) $result = $result();
        return $result;
    };
}
$factorial = trampoline(function(int $n, int $acc = 1) use (&$factorial): mixed {
    if ($n <= 1) return $acc;
    return fn() use ($n, $acc) => ($factorial)($n - 1, $n * $acc);
});
echo $factorial(5);

__vybe_check(ob_get_clean(), "120");
