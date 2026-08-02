<?php
// vybe-test: php/error_handling_deep/error_clear_last
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

set_error_handler(fn() => true);
trigger_error("something", E_USER_NOTICE);
restore_error_handler();
error_clear_last();
$err = error_get_last();
echo $err === null ? 'cleared' : 'still set';
