<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_boolean_conversion
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs
// vybe-test-mode: compile

$_POST["agree"] = "yes";
$boolVal = filter_input(INPUT_POST, "agree", FILTER_VALIDATE_BOOLEAN, FILTER_NULL_ON_FAILURE);
echo $boolVal ? "TRUE" : "FALSE";
