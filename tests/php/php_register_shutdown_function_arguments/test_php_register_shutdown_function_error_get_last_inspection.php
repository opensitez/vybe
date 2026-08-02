<?php
// vybe-test: php/php_register_shutdown_function_arguments/test_php_register_shutdown_function_error_get_last_inspection
// origin: languages/php/tests/php/test_php_register_shutdown_function_arguments.rs
// vybe-test-mode: compile

register_shutdown_function(function() {
    $err = error_get_last();
});
echo "ERROR_GET_LAST_SHUTDOWN_OK";
