<?php
// vybe-test: php/output_buffering/ob_list_handlers_before_and_after_start
// origin: languages/php/tests/php/test_output_buffering.rs

$before = ob_list_handlers();
ob_start();
$after = ob_list_handlers();
ob_end_clean();
echo (is_array($before) ? count($before) : 0) . '|' . (is_array($after) ? count($after) : 0);
