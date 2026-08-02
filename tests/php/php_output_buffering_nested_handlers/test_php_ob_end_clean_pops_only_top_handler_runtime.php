<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_end_clean_pops_only_top_handler_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start();
ob_start();
echo 'inner';
ob_end_clean();
echo 'outer';
$status = ob_get_status(true);
$n = is_array($status) ? count($status) : -1;
$contents = ob_get_contents();
ob_end_clean();
echo $contents . '|' . $n . '|' . $contents;
