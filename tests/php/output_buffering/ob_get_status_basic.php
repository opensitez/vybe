<?php
// vybe-test: php/output_buffering/ob_get_status_basic
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
echo "data";
$status = ob_get_status();
ob_end_clean();
echo isset($status['level']) ? 'has level' : 'no level';
