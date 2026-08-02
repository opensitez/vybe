<?php
// vybe-test: php/output_buffering/ob_start_with_empty_buffer_reports_zero_length
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
$len = ob_get_length();
ob_end_clean();
echo $len;
