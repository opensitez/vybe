<?php
// vybe-test: php/output_buffering/ob_get_length_reports_byte_count
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo '12345';
$len = ob_get_length();
ob_end_clean();
echo $len;
