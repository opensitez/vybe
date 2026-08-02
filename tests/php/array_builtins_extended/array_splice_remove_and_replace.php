<?php
// vybe-test: php/array_builtins_extended/array_splice_remove_and_replace
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["a", "b", "c", "d", "e"];
$removed = array_splice($a, 1, 2, ["x", "y", "z"]);
echo implode(",", $a);
echo implode(",", $removed);
