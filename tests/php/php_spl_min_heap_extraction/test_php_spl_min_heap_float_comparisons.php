<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_float_comparisons
// origin: languages/php/tests/php/test_php_spl_min_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
$heap->insert(3.14);
$heap->insert(1.41);
$heap->insert(2.71);
echo $heap->extract() === 1.41 ? "FLOAT_MIN_OK" : "FAIL";
