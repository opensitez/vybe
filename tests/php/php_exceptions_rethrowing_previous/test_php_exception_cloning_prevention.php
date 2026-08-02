<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_exception_cloning_prevention
// origin: languages/php/tests/php/test_php_exceptions_rethrowing_previous.rs
// vybe-test-mode: compile

$e1 = new Exception("Original");
try {
    $e2 = clone $e1;
} catch (Error $err) {
    echo "Cannot clone exception";
}
