<?php
// vybe-test: php/enum_advanced_runtime/enum_from_array_map
// origin: languages/php/tests/php/test_enum_advanced_runtime.rs

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

enum N: int { case A = 1; case B = 2; }
echo implode(',', array_map(fn(N $n) => $n->value, N::cases()));

__vybe_check(ob_get_clean(), "1,2");
