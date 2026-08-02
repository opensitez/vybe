<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_get_status_full_stack_array
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs

ob_start(); // Level 1
ob_start(); // Level 2
$status = ob_get_status(true);
ob_end_clean();
ob_end_clean();

echo "LevelsCount: " . count($status);
