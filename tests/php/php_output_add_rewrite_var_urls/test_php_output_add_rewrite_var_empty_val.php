<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_empty_val
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs
// vybe-test-mode: compile

if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("empty_var", "");
    output_reset_rewrite_vars();
}
echo "EMPTY_VAL_REWRITE_OK";
