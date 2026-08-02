<?php
// vybe-test: php/php_register_shutdown_function_arguments/test_php_register_shutdown_function_multiple_args
// origin: languages/php/tests/php/test_php_register_shutdown_function_arguments.rs
// vybe-test-mode: compile

register_shutdown_function(function($a, $b, $c, $d) {}, 1, "two", 3.0, true);
echo "MULTIPLE_ARGS_SHUTDOWN_OK";
