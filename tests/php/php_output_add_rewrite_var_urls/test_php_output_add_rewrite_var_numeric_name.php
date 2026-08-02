<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_numeric_name
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs
// vybe-test-mode: compile

if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("123", "val");
    output_reset_rewrite_vars();
}
echo "NUMERIC_NAME_REWRITE_OK";
