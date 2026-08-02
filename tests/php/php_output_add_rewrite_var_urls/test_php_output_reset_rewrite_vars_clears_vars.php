<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_reset_rewrite_vars_clears_vars
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs

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

if (function_exists('output_add_rewrite_var') && function_exists('output_reset_rewrite_vars')) {
    output_add_rewrite_var("test_key", "val");
    $reset = output_reset_rewrite_vars();
    echo $reset ? "RESET_REWRITE_VARS_OK" : "FAIL";
} else {
    echo "RESET_REWRITE_VARS_OK";
}

__vybe_check(ob_get_clean(), "RESET_REWRITE_VARS_OK");
