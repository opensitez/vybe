<?php
// vybe-test: php/union_types/union_type_float_or_int
// origin: languages/php/tests/php/test_union_types.rs

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

function toNum(int|float $n): string { return is_int($n) ? 'int' : 'float'; }
echo toNum(5) . ',' . toNum(5.5);

__vybe_check(ob_get_clean(), "int,float");
