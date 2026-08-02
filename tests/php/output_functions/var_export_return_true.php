<?php
// vybe-test: php/output_functions/var_export_return_true
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$data = ['x' => 1, 'y' => 'hello', 'z' => true];
$code = var_export($data, true);
echo is_string($code) ? 'string' : 'not string';
echo strlen($code) > 0 ? ':non-empty' : ':empty';
