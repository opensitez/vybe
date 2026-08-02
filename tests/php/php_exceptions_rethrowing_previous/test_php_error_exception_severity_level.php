<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_error_exception_severity_level
// origin: languages/php/tests/php/test_php_exceptions_rethrowing_previous.rs
// vybe-test-mode: compile

try {
    throw new ErrorException("Warning Exception", 0, E_WARNING, __FILE__, __LINE__);
} catch (ErrorException $e) {
    echo "Severity=" . $e->getSeverity();
}
