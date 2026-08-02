<?php
// vybe-test: php/advanced_closures/closure_modifies_external_array_by_ref
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$items = [];
$collect = function(mixed $v) use (&$items): void { $items[] = $v; };
$collect('a');
$collect('b');
$collect('c');
echo count($items);
