<?php
// vybe-test: php/output_buffering/ob_get_flush_with_binary_contents
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo "A";
ob_start();
echo "B";
$f = ob_get_flush();
$len = ob_get_length();
ob_end_clean();
echo $f . $len;
