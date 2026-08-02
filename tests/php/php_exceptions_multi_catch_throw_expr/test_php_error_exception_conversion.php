<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_error_exception_conversion
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs
// vybe-test-mode: compile

set_error_handler(function($severity, $message, $file, $line) {
    if (!(error_reporting() & $severity)) {
        return;
    }
    throw new ErrorException($message, 0, $severity, $file, $line);
});

try {
    // trigger user notice or warning
    trigger_error("User warning", E_USER_WARNING);
} catch (ErrorException $e) {
    echo "Converted error to exception: " . $e->getMessage();
} finally {
    restore_error_handler();
}
