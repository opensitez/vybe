<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_sanitize_encoded
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs
// vybe-test-mode: compile

$_GET["url"] = "https://example.com/test?a=1&b=2";
$clean = filter_input(INPUT_GET, "url", FILTER_SANITIZE_URL);
echo $clean;
