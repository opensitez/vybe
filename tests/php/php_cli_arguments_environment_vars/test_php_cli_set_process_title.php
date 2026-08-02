<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_cli_set_process_title
// origin: languages/php/tests/php/test_php_cli_arguments_environment_vars.rs
// vybe-test-mode: compile

if (function_exists('cli_set_process_title')) {
    @cli_set_process_title("vybe-worker");
    echo cli_get_process_title();
}
