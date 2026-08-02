<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_set_error_handler_e_all_mask
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs
// vybe-test-mode: compile

$count = 0;
set_error_handler(function() use (&$count) { $count++; return true; }, E_ALL);
@trigger_error("Notice 1", E_USER_NOTICE);
@trigger_error("Warning 1", E_USER_WARNING);
restore_error_handler();
echo $count === 2 ? "E_ALL_MASK_2_OK" : "FAIL";
