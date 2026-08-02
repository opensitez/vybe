<?php
// vybe-test: php/error_handling_deep/error_get_last_structure
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

set_error_handler(fn() => true);
trigger_error("structured error", E_USER_NOTICE);
restore_error_handler();
$err = error_get_last();
if ($err !== null) {
    echo isset($err['type'])    ? 'has type' : 'no type';
    echo isset($err['message']) ? ':has msg'  : ':no msg';
    echo isset($err['file'])    ? ':has file' : ':no file';
    echo isset($err['line'])    ? ':has line' : ':no line';
}
