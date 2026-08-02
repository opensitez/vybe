<?php
// vybe-test: php/php_intl_error_name_lookup/test_intl_is_failure_check
// origin: languages/php/tests/php/test_php_intl_error_name_lookup.rs

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

if (function_exists('intl_is_failure')) {
    echo intl_is_failure(0) ? 'fail' : 'success', "\n";
} else {
    echo "success\n";
}

__vybe_check(ob_get_clean(), "success");
