<?php
// vybe-test: php/php_spl_max_heap_extraction/test_php_spl_max_heap_count
// origin: languages/php/tests/php/test_php_spl_max_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMaxHeap();
$heap->insert(1);
$heap->insert(2);
$heap->insert(3);
echo count($heap) === 3 ? "COUNT_HEAP_OK" : "FAIL";
