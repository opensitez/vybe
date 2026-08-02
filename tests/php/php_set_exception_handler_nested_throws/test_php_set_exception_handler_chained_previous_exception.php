<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_set_exception_handler_chained_previous_exception
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs
// vybe-test-mode: compile

$hasPrevious = false;
set_exception_handler(function(Throwable $e) use (&$hasPrevious) {
    if ($e->getPrevious() !== null) $hasPrevious = true;
});
try {
    $first = new Exception("First");
    throw new Exception("Second", 0, $first);
} catch (Throwable $e) {
    $h = set_exception_handler(null);
    $h($e);
}
echo $hasPrevious ? "CHAINED_PREVIOUS_OK" : "FAIL";
