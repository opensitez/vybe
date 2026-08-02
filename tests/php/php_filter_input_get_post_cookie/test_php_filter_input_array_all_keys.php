<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_array_all_keys
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs
// vybe-test-mode: compile

$_GET["a"] = "10";
$_GET["b"] = "20";
$data = filter_input_array(INPUT_GET, FILTER_VALIDATE_INT);
echo is_array($data) ? "INPUT_ARRAY_OK" : "FAIL";
