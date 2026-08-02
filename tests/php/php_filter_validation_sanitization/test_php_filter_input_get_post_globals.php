<?php
// vybe-test: php/php_filter_validation_sanitization/test_php_filter_input_get_post_globals
// origin: languages/php/tests/php/test_php_filter_validation_sanitization.rs
// vybe-test-mode: compile

$_GET["id"] = "100";
$id = filter_input(INPUT_GET, "id", FILTER_VALIDATE_INT);
echo $id;
