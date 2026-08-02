<?php
// vybe-test: php/error_handling_deep/trigger_error_user_notice
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$msg = '';
set_error_handler(function(int $no, string $str) use (&$msg): bool { $msg = $str; return true; });
trigger_error("hello notice", E_USER_NOTICE);
restore_error_handler();
echo $msg;
