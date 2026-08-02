<?php
// vybe-test: php/output_functions/var_dump_nested
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$obj = new stdClass();
$obj->name = 'test';
$obj->values = [1, 2, 3];
ob_start();
var_dump($obj);
$output = ob_get_clean();
echo strlen($output) > 0 ? 'dumped' : 'empty';
