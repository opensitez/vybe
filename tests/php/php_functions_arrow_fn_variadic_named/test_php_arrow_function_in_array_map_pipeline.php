<?php
// vybe-test: php/php_functions_arrow_fn_variadic_named/test_php_arrow_function_in_array_map_pipeline
// origin: languages/php/tests/php/test_php_functions_arrow_fn_variadic_named.rs
// vybe-test-mode: compile

$users = [
    ["name" => "Alice", "active" => true],
    ["name" => "Bob", "active" => false],
    ["name" => "Charlie", "active" => true],
];

$activeNames = array_map(
    fn($u) => $u["name"],
    array_filter($users, fn($u) => $u["active"])
);

echo implode(",", $activeNames);
