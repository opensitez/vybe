<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_memory_get_usage_and_peak
// origin: languages/php/tests/php/test_php_cli_arguments_environment_vars.rs
// vybe-test-mode: compile

$alloc = memory_get_usage();
$peak = memory_get_peak_usage();
echo ($alloc > 0 && $peak >= $alloc) ? "MEM_USAGE_OK" : "MEM_FAIL";
