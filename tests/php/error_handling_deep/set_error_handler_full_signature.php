<?php
// vybe-test: php/error_handling_deep/set_error_handler_full_signature
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$last = [];
set_error_handler(function(int $errno, string $errstr, string $errfile, int $errline) use (&$last): bool {
    $last = ['no' => $errno, 'str' => $errstr, 'line' => $errline];
    return true;
});
trigger_error("custom error", E_USER_ERROR);
restore_error_handler();
echo $last['no'] === E_USER_ERROR ? 'correct errno' : 'wrong errno';
echo is_string($last['str']) ? ':has message' : ':no message';
