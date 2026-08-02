<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_get_level_for_nested_buffers_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start();
$l1 = ob_get_level();
ob_start();
$l2 = ob_get_level();
ob_end_clean();
$l3 = ob_get_level();
ob_end_clean();
echo $l1 . '|' . $l2 . '|' . $l3;
