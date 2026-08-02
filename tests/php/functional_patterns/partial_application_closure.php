<?php
// vybe-test: php/functional_patterns/partial_application_closure
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

function partial(callable $fn, mixed ...$partial): Closure {
    return function() use ($fn, $partial) {
        return $fn(...$partial, ...func_get_args());
    };
}
$multiply = fn($a,$b) => $a * $b;
$double = partial($multiply, 2);
$triple = partial($multiply, 3);
echo $double(7) . ',' . $triple(7);

__vybe_check(ob_get_clean(), "14,21");
