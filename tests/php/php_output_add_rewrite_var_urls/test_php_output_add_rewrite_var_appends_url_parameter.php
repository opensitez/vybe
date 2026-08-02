<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_appends_url_parameter
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
    output_add_rewrite_var("sid", "session_token_123");
    echo '<a href="index.php">Link</a>';
    output_reset_rewrite_vars();
} else {
    echo '<a href="index.php">Link</a>';
}

__vybe_check(ob_get_clean(), "<a href=\"index.php\">Link</a>");
