<?php
// vybe-test: php/php_output_add_rewrite_var_urls/test_php_output_add_rewrite_var_form_field_injection
// origin: languages/php/tests/php/test_php_output_add_rewrite_var_urls.rs
// vybe-test-mode: compile

if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("csrf", "token_abc");
    echo '<form action="post.php"><input type="text"/></form>';
    output_reset_rewrite_vars();
}
echo "FORM_REWRITE_CHECKED";
