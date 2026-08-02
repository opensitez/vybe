<?php
// vybe-test: php/output_buffering/ob_nested_buffers_flush_in_order
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'outer-';
ob_start();
echo 'inner';
$inner = ob_get_clean();
echo $inner . '-end';
$outer = ob_get_clean();
echo $outer;
