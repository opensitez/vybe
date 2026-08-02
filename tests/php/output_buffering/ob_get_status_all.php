<?php
// vybe-test: php/output_buffering/ob_get_status_all
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
ob_start();
$statuses = ob_get_status(true);
ob_end_clean();
ob_end_clean();
echo count($statuses) >= 2 ? 'two levels' : 'less';
