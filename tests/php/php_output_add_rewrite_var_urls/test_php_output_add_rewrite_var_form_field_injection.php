<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_form_field_injection
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
    output_add_rewrite_var("csrf", "token_abc");
    echo '<form action="post.php"><input type="text"/></form>';
    output_reset_rewrite_vars();
}
echo "FORM_REWRITE_CHECKED";


__vybe_check(ob_get_clean(), "<form action=\"post.php\"><input type=\"text\"/></form>FORM_REWRITE_CHECKED");
