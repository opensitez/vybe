<?php
// vybe-test: php/object_comparison/spl_object_id_same_for_same_reference
// origin: languages/php/tests/php/test_object_comparison.rs

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

class A {}
$a = new A(); $b = $a;
echo spl_object_id($a) === spl_object_id($b) ? 'same' : 'different';

__vybe_check(ob_get_clean(), "same");
