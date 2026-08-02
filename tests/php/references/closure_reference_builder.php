<?php
// vybe-test: php/references/closure_reference_builder
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$result = [];
$collect = function($v) use (&$result) { $result[] = $v * $v; };
array_map($collect, [1, 2, 3]);
echo implode(',', $result);
