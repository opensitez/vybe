<?php
// vybe-test: php/filter_has_var_superglobals/filter_has_var_basic
// origin: languages/php/tests/php/test_filter_has_var_superglobals.rs

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

// We test if the function exists and can handle empty superglobals gracefully
echo filter_has_var(INPUT_GET, 'test') ? "yes" : "no";

__vybe_check(ob_get_clean(), "no");
