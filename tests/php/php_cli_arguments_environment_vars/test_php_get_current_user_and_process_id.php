<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_get_current_user_and_process_id
// origin: languages/php/tests/php/test_php_cli_arguments_environment_vars.rs
// vybe-test-mode: compile

$pid = getmypid();
$user = get_current_user();
echo "PID=$pid USER=$user";
