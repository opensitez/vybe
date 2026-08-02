<?php
// vybe-test: php/error_handling_deep/trigger_error_user_deprecated
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$caught = false;
set_error_handler(function(int $no) use (&$caught): bool {
    if ($no === E_USER_DEPRECATED) $caught = true;
    return true;
});
trigger_error("use newFunc() instead", E_USER_DEPRECATED);
restore_error_handler();
echo $caught ? 'deprecated caught' : 'missed';
