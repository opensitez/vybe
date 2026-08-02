<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_get_length_after_nested_clean_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start();
ob_start();
echo 'inner';
$inner_len = ob_get_length();
ob_end_clean();
$outer_len = ob_get_length();
ob_end_clean();
echo $inner_len . '|' . $outer_len;
