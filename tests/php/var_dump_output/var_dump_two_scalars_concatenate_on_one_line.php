<?php
// vybe-test: php/var_dump_output/var_dump_two_scalars_concatenate_on_one_line
// origin: languages/php/tests/php/test_var_dump_output.rs

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

var_dump('a');
var_dump('b');

__vybe_check(ob_get_clean(), "string(1) \"a\"\nstring(1) \"b\"");
