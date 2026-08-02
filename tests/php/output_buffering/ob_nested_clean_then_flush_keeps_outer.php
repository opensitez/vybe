<?php
// vybe-test: php/output_buffering/ob_nested_clean_then_flush_keeps_outer
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'outer-';
ob_start(fn(string $buf): string => '[' . $buf . ']');
echo 'inner';
ob_end_flush();
$outer = ob_get_clean();
echo $outer;
