<?php
// vybe-test: php/output_buffering/ob_nested_two_levels
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
echo "outer";
ob_start();
echo "inner";
$inner = ob_get_clean();
echo "-" . $inner . "-";
$outer = ob_get_clean();
echo $outer;
