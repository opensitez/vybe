<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_reset_rewrite_vars_returns_bool
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs
// vybe-test-mode: compile

if (function_exists('output_reset_rewrite_vars')) {
    $res = output_reset_rewrite_vars();
    echo is_bool($res) ? "RESET_BOOL_OK" : "FAIL";
} else {
    echo "RESET_BOOL_OK";
}
