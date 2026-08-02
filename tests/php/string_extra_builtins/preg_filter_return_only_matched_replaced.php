<?php
// vybe-test: php/string_extra_builtins/preg_filter_return_only_matched_replaced
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$input = ["foo1", "bar", "foo2", "baz", "foo3"];
$result = preg_filter('/^foo(\d)/', 'match$1', $input);
echo count($result);
echo is_array($result) ? "array" : "not";
