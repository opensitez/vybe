<?php
// vybe-test: php/string_interpolation_complex/interpolation_newline_escape
// origin: languages/php/tests/php/test_string_interpolation_complex.rs

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

$s = "line1\nline2";
echo substr_count($s, "\n");
echo "\n";

__vybe_check(ob_get_clean(), "1");
