<?php
// vybe-test: php/union_types_runtime/union_property_set_multiple_types
// origin: languages/php/tests/php/test_union_types_runtime.rs

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

class Holder { public int|float $n = 1; }
$h = new Holder();
$h->n = 2.5;
echo (string)$h->n;

__vybe_check(ob_get_clean(), "2.5");
