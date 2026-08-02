<?php
// vybe-test: php/class_introspection/get_class_methods_lists_public_method
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

class M { public function go(): void {} }
echo in_array('go', get_class_methods('M'), true) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes");
