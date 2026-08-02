<?php
// vybe-test: php/output_buffering/ob_clean_then_get_length_zero
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'temp';
ob_clean();
$r = ob_get_length() === 0 ? 'zero' : 'not';
ob_end_clean();
echo $r;
