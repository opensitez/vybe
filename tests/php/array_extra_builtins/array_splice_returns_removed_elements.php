<?php
// vybe-test: php/array_extra_builtins/array_splice_returns_removed_elements
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = ["a", "b", "c", "d", "e"];
$removed = array_splice($a, 1, 2);
echo count($removed);
echo implode(",", $removed);
echo count($a);
