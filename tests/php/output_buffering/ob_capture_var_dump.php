<?php
// vybe-test: php/output_buffering/ob_capture_var_dump
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
var_dump(42, 'hello', true);
$output = ob_get_clean();
echo strlen($output) > 0 ? 'ok' : 'fail';
