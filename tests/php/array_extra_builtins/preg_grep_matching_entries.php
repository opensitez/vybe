<?php
// vybe-test: php/array_extra_builtins/preg_grep_matching_entries
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$input = ["foo1", "bar", "foo2", "baz", "foo3"];
$matches = preg_grep('/^foo/', $input);
echo count($matches);
echo is_array($matches) ? "array" : "not";
