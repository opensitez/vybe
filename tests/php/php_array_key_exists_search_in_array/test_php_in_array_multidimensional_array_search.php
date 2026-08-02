<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_in_array_multidimensional_array_search
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs
// vybe-test-mode: compile

$user1 = ["id" => 1, "name" => "Alice"];
$user2 = ["id" => 2, "name" => "Bob"];
$list = [$user1, $user2];
echo in_array(["id" => 1, "name" => "Alice"], $list, true) ? "STRICT_STRUCT_FOUND" : "FAIL";
