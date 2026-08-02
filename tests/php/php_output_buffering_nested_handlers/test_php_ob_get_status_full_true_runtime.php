<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_get_status_full_true_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start();
echo "x";
$status = ob_get_status(true);
ob_end_clean();
echo is_array($status) ? (is_array(array_values($status)[0]) ? 'full' : 'partial') : 'false';
