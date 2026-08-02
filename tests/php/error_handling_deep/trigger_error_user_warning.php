<?php
// vybe-test: php/error_handling_deep/trigger_error_user_warning
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$caught = false;
set_error_handler(function() use (&$caught): bool { $caught = true; return true; });
trigger_error("test", E_USER_WARNING);
restore_error_handler();
echo $caught ? 'triggered' : 'not triggered';
