<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_argc_argv_global_variables
// origin: languages/php/tests/php/test_php_cli_arguments_environment_vars.rs
// vybe-test-mode: compile

global $argc, $argv;
echo "Args count: " . (is_array($argv) ? count($argv) : 0);
