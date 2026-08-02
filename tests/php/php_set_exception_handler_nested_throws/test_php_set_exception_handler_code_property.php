<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_code_property
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs
// vybe-test-mode: compile

$codeVal = 0;
set_exception_handler(function(Throwable $e) use (&$codeVal) {
    $codeVal = $e->getCode();
});
try {
    throw new Exception("With code", 404);
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $codeVal === 404 ? "EXCEPTION_CODE_404_OK" : "FAIL";
