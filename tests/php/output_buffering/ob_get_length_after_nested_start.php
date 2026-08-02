<?php
// vybe-test: php/output_buffering/ob_get_length_after_nested_start
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_start();
echo 'nested';
$len = ob_get_length();
ob_end_clean();
ob_end_clean();
echo $len;
