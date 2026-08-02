<?php
// vybe-test: php/array_builtins_extended/array_column_extract_field
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$rows = [
    ["id" => 1, "name" => "Alice", "dept" => "Eng"],
    ["id" => 2, "name" => "Bob",   "dept" => "Mkt"],
    ["id" => 3, "name" => "Carol", "dept" => "Eng"],
];
$names = array_column($rows, "name");
echo implode(",", $names);
$byId = array_column($rows, "dept", "id");
echo $byId[2];
