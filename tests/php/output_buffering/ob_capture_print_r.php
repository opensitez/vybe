<?php
// vybe-test: php/output_buffering/ob_capture_print_r
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

$data = ['a' => 1, 'b' => 2, 'c' => 3];
ob_start();
print_r($data);
$output = ob_get_clean();
echo strlen($output) > 0 ? 'captured' : 'empty';
