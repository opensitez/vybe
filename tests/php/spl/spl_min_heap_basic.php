<?php
// vybe-test: php/spl/spl_min_heap_basic
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
$heap->insert(5);
$heap->insert(2);
$heap->insert(8);
$heap->insert(1);
$result = [];
while (!$heap->isEmpty()) { $result[] = $heap->extract(); }
echo implode(',', $result);
