<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_multiple_key_value_pairs
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs
// vybe-test-mode: compile

if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("k1", "v1");
    output_add_rewrite_var("k2", "v2");
    output_reset_rewrite_vars();
}
echo "MULTIPLE_REWRITE_VARS_OK";
