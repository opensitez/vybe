<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_count_after_extractions
// origin: languages/php/tests/php/test_php_spl_min_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
$heap->insert(100);
$heap->insert(200);
$heap->extract();
echo count($heap) === 1 ? "COUNT_1_OK" : "FAIL";
