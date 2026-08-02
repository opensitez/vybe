<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_throwable_type_hint_error
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs
// vybe-test-mode: compile

$caughtError = false;
set_exception_handler(function(Throwable $t) use (&$caughtError) {
    if ($t instanceof TypeError) $caughtError = true;
});
try {
    throw new TypeError("Type error thrown");
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $caughtError ? "TYPE_ERROR_CAUGHT_OK" : "FAIL";
