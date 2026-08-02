<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_set_error_handler_deprecated_level
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs
// vybe-test-mode: compile

$dep = false;
set_error_handler(function($errno) use (&$dep) {
    if ($errno === E_USER_DEPRECATED) $dep = true;
    return true;
});
@trigger_error("Deprecated feature used", E_USER_DEPRECATED);
restore_error_handler();
echo $dep ? "E_USER_DEPRECATED_OK" : "FAIL";
