<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_set_error_handler_class_method_array_callable
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs
// vybe-test-mode: compile

class ErrorCatcher {
    public static function handle($errno, $errstr): bool { return true; }
}
set_error_handler([ErrorCatcher::class, "handle"]);
@trigger_error("Class handler test", E_USER_NOTICE);
restore_error_handler();
echo "CLASS_METHOD_HANDLER_OK";
