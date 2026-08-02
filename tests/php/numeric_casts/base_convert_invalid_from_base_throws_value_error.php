<?php
// vybe-test: php/numeric_casts/base_convert_invalid_from_base_throws_value_error
// origin: languages/php/tests/php/test_numeric_casts.rs

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

try { base_convert('10', 1, 10); echo 'ok'; }
catch (ValueError $e) { echo 'base'; }

__vybe_check(ob_get_clean(), "base");
