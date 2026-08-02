<?php
// vybe-test: php/output_buffering/ob_nested_levels_with_midlevel_clean
// origin: languages/php/tests/php/test_output_buffering.rs

echo 'L0-';
ob_start();
echo 'L1';
ob_start();
echo 'L2';
$l2 = ob_get_clean();
echo ':' . $l2 . ':';
$l1 = ob_get_clean();
echo $l1;
