<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_getenv_local_array_parsing
// origin: languages/php/tests/php/test_php_cli_arguments_environment_vars.rs
// vybe-test-mode: compile

$env = getenv();
echo is_array($env) ? "ENV_ARRAY" : "FAIL";
