<?php
// vybe-test: php/function_builtins/array_walk_with_key_and_extra
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

$fruits = ['a' => 'apple', 'b' => 'banana', 'c' => 'cherry'];
array_walk($fruits, function(&$value, $key, $prefix) {
    $value = $prefix . ':' . $key . '=' . $value;
}, 'fruit');
echo implode(',', $fruits);
