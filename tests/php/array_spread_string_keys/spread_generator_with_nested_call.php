<?php
// vybe-test: php/array_spread_string_keys/spread_generator_with_nested_call
// origin: languages/php/tests/php/test_array_spread_string_keys.rs

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

function gen(): Generator { yield from [1, 2, 3]; }
function sum3(int $a, int $b, int $c): int { return $a + $b + $c; }
echo sum3(...gen());

__vybe_check(ob_get_clean(), "6");
