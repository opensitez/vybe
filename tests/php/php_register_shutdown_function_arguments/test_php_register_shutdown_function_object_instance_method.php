<?php
// vybe-test: php/php_register_shutdown_function_arguments/test_php_register_shutdown_function_object_instance_method
// origin: languages/php/tests/php/test_php_register_shutdown_function_arguments.rs
// vybe-test-mode: compile

class Logger {
    public function flush() {}
}
$log = new Logger();
register_shutdown_function([$log, "flush"]);
echo "OBJECT_INSTANCE_SHUTDOWN_OK";
