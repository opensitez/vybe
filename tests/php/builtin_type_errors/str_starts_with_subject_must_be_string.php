<?php
// vybe-test: php/builtin_type_errors/str_starts_with_subject_must_be_string
// origin: languages/php/tests/php/test_builtin_type_errors.rs

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

try { str_starts_with([], 'a'); echo 'ok'; }
catch (TypeError $e) { echo 'starts'; }

__vybe_check(ob_get_clean(), "starts");
