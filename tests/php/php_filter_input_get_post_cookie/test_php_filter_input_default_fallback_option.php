<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_default_fallback_option
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs
// vybe-test-mode: compile

$val = filter_input(INPUT_GET, "missing_key", FILTER_VALIDATE_INT, [
    "options" => ["default" => 100]
]);
echo "Fallback: $val";
