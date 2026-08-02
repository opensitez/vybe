<?php
// vybe-test: php/php_register_shutdown_function_arguments/test_php_register_shutdown_function_class_static_method
// origin: languages/php/tests/php/test_php_register_shutdown_function_arguments.rs
// vybe-test-mode: compile

class Cleaner {
    public static function cleanUp($target) {}
}
register_shutdown_function([Cleaner::class, "cleanUp"], "temp_folder");
echo "CLASS_STATIC_SHUTDOWN_OK";
