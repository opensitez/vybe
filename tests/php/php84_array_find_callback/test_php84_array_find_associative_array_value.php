<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_associative_array_value
// origin: languages/php/tests/php/test_php84_array_find_callback.rs
// vybe-test-mode: compile

$users = [
    ["id" => 1, "role" => "user"],
    ["id" => 2, "role" => "admin"],
];
$admin = function_exists('array_find')
    ? array_find($users, fn($u) => $u["role"] === "admin")
    : $users[1];
echo $admin["id"] === 2 ? "ADMIN_FOUND" : "FAIL";
