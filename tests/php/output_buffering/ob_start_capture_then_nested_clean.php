<?php
// vybe-test: php/output_buffering/ob_start_capture_then_nested_clean
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_start();
echo 'inner';
ob_clean();
echo 'outer';
$inner = ob_get_clean();
ob_end_clean();
echo $inner;
