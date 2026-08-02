<?php
// vybe-test: php/error_handling_deep/set_error_handler_with_level
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$caught = 0;
set_error_handler(function() use (&$caught): bool { $caught++; return true; }, E_USER_NOTICE);
trigger_error("a notice", E_USER_NOTICE);
trigger_error("a warning", E_USER_WARNING); // not caught by this handler
restore_error_handler();
echo $caught;
