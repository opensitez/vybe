<?php
// vybe-test: php/class_introspection/get_object_vars_public_only
// origin: languages/php/tests/php/test_class_introspection.rs

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

class V { public int $a = 1; private int $b = 2; }
echo json_encode(get_object_vars(new V()));

__vybe_check(ob_get_clean(), "{\"a\":1}");
