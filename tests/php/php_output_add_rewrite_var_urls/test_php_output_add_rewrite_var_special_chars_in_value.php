<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_special_chars_in_value
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs
// vybe-test-mode: compile

if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("tag", "a & b = c");
    output_reset_rewrite_vars();
}
echo "SPECIAL_CHARS_REWRITE_OK";
