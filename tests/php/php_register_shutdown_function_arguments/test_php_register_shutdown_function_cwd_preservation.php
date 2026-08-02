<?php
// vybe-test: php/php_register_shutdown_function_arguments/test_php_register_shutdown_function_cwd_preservation
// origin: languages/php/tests/php/test_php_register_shutdown_function_arguments.rs
// vybe-test-mode: compile

$cwd = getcwd();
register_shutdown_function(function() use ($cwd) {
    // In shutdown functions, working directory might change to root depending on SAPI
});
echo "CWD_SHUTDOWN_OK";
