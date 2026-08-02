<?php
// vybe-test: php/error_handling_deep/error_get_last_basic
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

set_error_handler(fn() => true); // suppress
@trigger_error("test error", E_USER_WARNING);
restore_error_handler();
$err = error_get_last();
echo $err !== null ? 'has error' : 'no error';
