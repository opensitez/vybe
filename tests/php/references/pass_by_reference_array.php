<?php
// vybe-test: php/references/pass_by_reference_array
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

function addItem(array &$arr, $item) { $arr[] = $item; }
$list = [1, 2];
addItem($list, 3);
echo count($list);
