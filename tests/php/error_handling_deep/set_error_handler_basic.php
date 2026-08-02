<?php
// vybe-test: php/error_handling_deep/set_error_handler_basic
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$errors = [];
set_error_handler(function(int $errno, string $errstr) use (&$errors): bool {
    $errors[] = "$errno: $errstr";
    return true; // suppress default handler
});
trigger_error("test warning", E_USER_WARNING);
restore_error_handler();
echo count($errors) > 0 ? 'caught' : 'missed';
