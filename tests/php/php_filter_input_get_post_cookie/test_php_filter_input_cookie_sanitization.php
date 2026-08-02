<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_cookie_sanitization
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs
// vybe-test-mode: compile

$_COOKIE["user_id"] = "12345";
$id = filter_input(INPUT_COOKIE, "user_id", FILTER_VALIDATE_INT);
echo "User ID: $id";
