<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_trace_as_string_inspection
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs
// vybe-test-mode: compile

$traceCaptured = false;
set_exception_handler(function(Throwable $e) use (&$traceCaptured) {
    if (strlen($e->getTraceAsString()) > 0) $traceCaptured = true;
});
try {
    throw new Exception("Trace check");
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $traceCaptured ? "TRACE_AS_STRING_OK" : "FAIL";
