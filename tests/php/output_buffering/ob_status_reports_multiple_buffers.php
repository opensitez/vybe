<?php
// vybe-test: php/output_buffering/ob_status_reports_multiple_buffers
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_start();
$status = ob_get_status(true);
$ok = is_array($status) && array_key_exists(0, $status) && is_array($status[0]);
ob_end_clean();
ob_end_clean();
echo $ok ? 'one' : 'none';
