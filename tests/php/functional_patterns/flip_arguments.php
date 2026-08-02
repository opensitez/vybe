<?php
// vybe-test: php/functional_patterns/flip_arguments
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

function flip(callable $fn): Closure {
    return fn($a,$b) => $fn($b,$a);
}
$sub = fn($a,$b) => $a - $b;
echo flip($sub)(3, 10);

__vybe_check(ob_get_clean(), "7");
