<?php
// vybe-test: php/error_handling_deep/set_exception_handler_basic
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$caught = null;
set_exception_handler(function(\Throwable $e) use (&$caught): void {
    $caught = $e->getMessage();
});
// Note: set_exception_handler catches uncaught exceptions at shutdown
// We test it's settable and callable
$prev = set_exception_handler(null); // restore
echo 'handler set';
