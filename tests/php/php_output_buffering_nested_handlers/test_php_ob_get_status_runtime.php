<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_get_status_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start();
echo "abc";
$status = ob_get_status();
ob_end_clean();
echo $status['name'] . '|' . $status['level'];
