<?php
// vybe-test: php/php84/array_find_key_basic
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

$items = ['apple' => 1.5, 'banana' => 0.5, 'cherry' => 3.0];
$key = array_find_key($items, fn($price) => $price > 2.0);
echo $key;
