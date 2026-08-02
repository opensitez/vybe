<?php
// vybe-test: php/destructure_assignment/list_destructure_typed_property_mismatch_throws
// origin: languages/php/tests/php/test_destructure_assignment.rs

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

class Holder { public int $n; }
$h = new Holder();
try { [$h->n] = ['not-int']; echo 'ok'; }
catch (TypeError $e) { echo 'typed'; }

__vybe_check(ob_get_clean(), "typed");
