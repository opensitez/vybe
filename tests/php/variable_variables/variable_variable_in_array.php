<?php
// vybe-test: php/variable_variables/variable_variable_in_array
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

$fields = ['name', 'age', 'city'];
$name = 'Alice';
$age  = 30;
$city = 'NY';
$out = [];
foreach ($fields as $f) { $out[] = $$f; }
echo implode(',', $out);
