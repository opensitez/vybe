<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_set_error_handler_stack_restore
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs
// vybe-test-mode: compile

$h1 = fn() => true;
$h2 = fn() => true;
set_error_handler($h1);
set_error_handler($h2);
$prev = restore_error_handler();
restore_error_handler();
echo "ERROR_HANDLER_STACK_OK";
