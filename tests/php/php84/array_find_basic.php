<?php
// vybe-test: php/php84/array_find_basic
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

$users = [
    ['name' => 'Alice', 'age' => 28],
    ['name' => 'Bob',   'age' => 35],
    ['name' => 'Carol', 'age' => 22],
];
$found = array_find($users, fn($u) => $u['age'] > 30);
echo $found['name'];
