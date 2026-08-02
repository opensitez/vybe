<?php
// vybe-test: php/php_spl_max_heap_extraction/test_php_spl_max_heap_duplicate_values
// origin: languages/php/tests/php/test_php_spl_max_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMaxHeap();
$heap->insert(42);
$heap->insert(42);
$heap->insert(10);
echo $heap->extract() === 42 && $heap->extract() === 42 ? "DUPLICATE_HEAPS_OK" : "FAIL";
