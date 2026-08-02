<?php
// vybe-test: php/catch_type_union_order/catch_concrete_class_not_interface_alone_on_unrelated
// origin: languages/php/tests/php/test_catch_type_union_order.rs

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

interface Marker {}
class Plain extends Exception {}
try { throw new Plain('plain'); }
catch (Marker $m) { echo 'marker'; }
catch (Exception $e) { echo 'exception'; }

__vybe_check(ob_get_clean(), "exception");
