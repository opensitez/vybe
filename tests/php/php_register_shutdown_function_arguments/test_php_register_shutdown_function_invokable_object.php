<?php
// vybe-test: php/php_register_shutdown_function_arguments/test_php_register_shutdown_function_invokable_object
// origin: languages/php/tests/php/test_php_register_shutdown_function_arguments.rs
// vybe-test-mode: compile

class ShutdownTask {
    public function __invoke($msg) {}
}
register_shutdown_function(new ShutdownTask(), "task_done");
echo "INVOKABLE_SHUTDOWN_OK";
