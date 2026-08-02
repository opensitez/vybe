<?php
// vybe-test: php/output_functions/print_r_return_true
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$data = ['a' => 1, 'b' => [2, 3]];
$output = print_r($data, true);
echo is_string($output) ? 'string' : 'not string';
echo strlen($output) > 0 ? ':non-empty' : ':empty';
