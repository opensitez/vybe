<?php
// vybe-test: php/output_buffering/ob_implicit_flush_false_retains_buffered_output
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_implicit_flush(false);
echo 'x';
$len = ob_get_length();
ob_end_clean();
echo $len . '|done';
