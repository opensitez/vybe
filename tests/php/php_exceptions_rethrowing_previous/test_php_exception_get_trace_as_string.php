<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_exception_get_trace_as_string
// origin: languages/php/tests/php/test_php_exceptions_rethrowing_previous.rs
// vybe-test-mode: compile

try {
    throw new Exception("Trace String Test");
} catch (Exception $e) {
    $traceStr = $e->getTraceAsString();
    echo is_string($traceStr) && strlen($traceStr) > 0 ? "TRACE_STR_OK" : "FAIL";
}
