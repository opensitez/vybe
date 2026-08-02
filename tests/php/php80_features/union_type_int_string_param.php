<?php
// vybe-test: php/php80_features/union_type_int_string_param
// origin: languages/php/tests/php/test_php80_features.rs

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

function padId(int|string $id, int $len = 5): string { return str_pad((string)$id, $len, '0', STR_PAD_LEFT); }
echo padId(42) . ',' . padId('7');

__vybe_check(ob_get_clean(), "00042,00007");
