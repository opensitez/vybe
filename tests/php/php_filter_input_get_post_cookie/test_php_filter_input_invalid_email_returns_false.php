<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_invalid_email_returns_false
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs
// vybe-test-mode: compile

$_GET["email"] = "not_an_email";
$res = filter_input(INPUT_GET, "email", FILTER_VALIDATE_EMAIL);
echo $res === false ? "EMAIL_INVALID" : "VALID";
