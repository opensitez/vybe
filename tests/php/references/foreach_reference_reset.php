<?php
// vybe-test: php/references/foreach_reference_reset
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$items = ['a', 'b', 'c'];
foreach ($items as &$item) { $item = strtoupper($item); }
unset($item);
echo implode('', $items);
