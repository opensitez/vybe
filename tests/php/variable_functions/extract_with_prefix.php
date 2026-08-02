<?php
// vybe-test: php/variable_functions/extract_with_prefix
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$name = 'outer';
$data = ['name' => 'inner', 'age' => 25];
extract($data, EXTR_PREFIX_ALL, 'row');
echo $name;
echo $row_name;
echo $row_age;
