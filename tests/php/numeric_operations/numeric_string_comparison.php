<?php
// vybe-test: php/numeric_operations/numeric_string_comparison
// origin: languages/php/tests/php/test_numeric_operations.rs

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

var_dump('1' == 1);
var_dump('01' == '1');
var_dump('10' == '1e1');
var_dump('0' == false);
var_dump('0' === false);

__vybe_check(ob_get_clean(), "bool(true)\nbool(true)\nbool(true)\nbool(true)\nbool(false)");
