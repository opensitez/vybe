<?php
// vybe-test: php/unserialize_errors/unserialize_unknown_class_becomes_incomplete_class
// origin: languages/php/tests/php/test_unserialize_errors.rs

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

$v = unserialize('O:15:"Missing\\Klass":0:{}');
echo is_object($v) && $v instanceof __PHP_Incomplete_Class ? 'incomplete' : 'other';

__vybe_check(ob_get_clean(), "incomplete");
