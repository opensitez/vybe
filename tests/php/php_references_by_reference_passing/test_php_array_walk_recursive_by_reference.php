<?php
// vybe-test: php/php_references_by_reference_passing/test_php_array_walk_recursive_by_reference
// origin: languages/php/tests/php/test_php_references_by_reference_passing.rs
// vybe-test-mode: compile

$nested = ["a" => 1, "b" => [2, 3]];
array_walk_recursive($nested, function(&$val) {
    $val += 10;
});
echo $nested["a"] . " " . $nested["b"][0];
