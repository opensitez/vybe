<?php
// vybe-test: php/php_register_shutdown_function_arguments/test_php_register_shutdown_function_string_function_name
// origin: languages/php/tests/php/test_php_register_shutdown_function_arguments.rs
// vybe-test-mode: compile

function myShutdownFunc() {}
register_shutdown_function("myShutdownFunc");
echo "STRING_NAME_SHUTDOWN_OK";
