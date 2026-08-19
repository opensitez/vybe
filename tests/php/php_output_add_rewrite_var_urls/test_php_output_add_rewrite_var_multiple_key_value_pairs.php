<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_multiple_key_value_pairs
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

if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("k1", "v1");
    output_add_rewrite_var("k2", "v2");
    output_reset_rewrite_vars();
}
echo "MULTIPLE_REWRITE_VARS_OK";


__vybe_check(ob_get_clean(), "MULTIPLE_REWRITE_VARS_OK");
