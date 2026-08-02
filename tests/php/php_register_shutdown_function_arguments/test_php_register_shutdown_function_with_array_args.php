<?php
// vybe-test: php/php_register_shutdown_function_arguments/test_php_register_shutdown_function_with_array_args
// origin: languages/php/tests/php/test_php_register_shutdown_function_arguments.rs
// vybe-test-mode: compile

register_shutdown_function(function($arr) {
    //
}, ["a" => 1, "b" => 2]);
echo "ARRAY_ARG_SHUTDOWN_OK";
