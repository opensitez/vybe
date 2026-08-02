<?php
// vybe-test: php/array_builtins_extended/extract_variables_from_array
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$data = ["username" => "alice", "role" => "admin", "level" => 5];
extract($data);
echo $username;
echo $role;
echo $level;
