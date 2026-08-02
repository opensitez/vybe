<?php
// vybe-test: php/output_buffering/ob_nested_buffers_report_level_changes
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'A';
ob_start();
echo 'B';
$l2 = ob_get_level();
ob_get_clean();
$l1 = ob_get_level();
ob_end_clean();
echo $l2 . '|' . $l1;
