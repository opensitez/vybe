<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_exception_to_string_formatting
// origin: languages/php/tests/php/test_php_exceptions_rethrowing_previous.rs
// vybe-test-mode: compile

try {
    throw new Exception("Test stringification", 500);
} catch (Exception $e) {
    $str = (string)$e;
    echo str_contains($str, "Test stringification") ? "STRINGIFIED_OK" : "FAIL";
}
