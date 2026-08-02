<?php
// vybe-test: php/arrays/array_unpack_default_overrides
// origin: languages/php/tests/php/test_arrays.rs

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

$base = [1, 2];
$merged = [...$base, ...[2 => 30, 3 => 40]];
echo count($merged) . "\n";
echo $merged[0] . ',' . $merged[2] . ',' . $merged[3];

__vybe_check(ob_get_clean(), "4\n1,30,40");
