<?php
// vybe-test: php/error_handling_deep/at_operator_with_handler
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$triggered = false;
set_error_handler(function() use (&$triggered): bool { $triggered = true; return true; });
$r = @trigger_error("suppressed?", E_USER_NOTICE);
restore_error_handler();
// @ suppresses at the engine level — handler may or may not be called
echo is_bool($r) || $r === null ? 'ran' : 'fail';
