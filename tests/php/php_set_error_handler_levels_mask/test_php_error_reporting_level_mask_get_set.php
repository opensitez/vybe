<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_error_reporting_level_mask_get_set
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs
// vybe-test-mode: compile

$old = error_reporting(E_ALL & ~E_NOTICE);
$current = error_reporting();
error_reporting($old);
echo $current === (E_ALL & ~E_NOTICE) ? "ERROR_REPORTING_MASK_OK" : "FAIL";
