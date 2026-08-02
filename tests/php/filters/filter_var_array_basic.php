<?php
// vybe-test: php/filters/filter_var_array_basic
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$data = [
    'age'   => '25',
    'email' => 'user@example.com',
    'score' => '3.14',
];
$filters = [
    'age'   => FILTER_VALIDATE_INT,
    'email' => FILTER_VALIDATE_EMAIL,
    'score' => FILTER_VALIDATE_FLOAT,
];
$result = filter_var_array($data, $filters);
var_dump($result['age']);
echo $result['email'] !== false ? 'valid email' : 'invalid email';
