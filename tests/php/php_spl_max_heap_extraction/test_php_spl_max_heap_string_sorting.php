<?php
// vybe-test: php/php_spl_max_heap_extraction/test_php_spl_max_heap_string_sorting
// origin: languages/php/tests/php/test_php_spl_max_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMaxHeap();
$heap->insert("apple");
$heap->insert("zebra");
$heap->insert("banana");
echo $heap->extract() === "zebra" ? "STRING_MAX_HEAP_OK" : "FAIL";
