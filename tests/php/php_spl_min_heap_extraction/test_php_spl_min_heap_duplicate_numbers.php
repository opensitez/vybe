<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_duplicate_numbers
// origin: languages/php/tests/php/test_php_spl_min_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
$heap->insert(7);
$heap->insert(7);
$heap->insert(2);
echo $heap->extract() === 2 && $heap->extract() === 7 ? "MIN_DUPLICATES_OK" : "FAIL";
