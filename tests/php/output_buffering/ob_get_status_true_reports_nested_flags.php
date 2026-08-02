<?php
// vybe-test: php/output_buffering/ob_get_status_true_reports_nested_flags
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_start();
$status = ob_get_status(true);
ob_end_clean();
ob_end_clean();
echo is_array($status) && count($status) >= 2 ? 'yes' : 'no';
