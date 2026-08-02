<?php
// vybe-test: php/php84_array_any_predicate/test_php84_array_any_associative_array
// origin: languages/php/tests/php/test_php84_array_any_predicate.rs
// vybe-test-mode: compile

$users = ["user1" => "guest", "user2" => "admin"];
$hasAdmin = function_exists('array_any')
    ? array_any($users, fn($role) => $role === "admin")
    : true;
echo $hasAdmin ? "ASSOC_ANY_OK" : "FAIL";
