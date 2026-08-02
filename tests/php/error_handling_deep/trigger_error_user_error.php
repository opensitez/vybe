<?php
// vybe-test: php/error_handling_deep/trigger_error_user_error
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$caught = false;
set_error_handler(function() use (&$caught): bool { $caught = true; return true; });
trigger_error("fatal-like error", E_USER_ERROR);
restore_error_handler();
echo $caught ? 'caught' : 'missed';
