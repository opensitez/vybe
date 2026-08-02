<?php
// vybe-test: php/output_buffering/ob_get_length_after_nested_operations
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'abc';
ob_start();
echo 'de';
echo $inner_len = ob_get_length();
ob_end_flush();
echo '|';
echo $outer_len = ob_get_length();
echo '|';
echo ob_get_clean();
