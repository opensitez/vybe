<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_ob_get_level_and_ob_get_status
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs
// vybe-test-mode: compile

$initialLevel = ob_get_level();
ob_start();
echo "Level active: " . (ob_get_level() - $initialLevel);
$status = ob_get_status();
print_r($status);
ob_end_clean();
