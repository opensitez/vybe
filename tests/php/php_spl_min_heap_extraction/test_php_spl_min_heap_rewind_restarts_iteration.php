<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_rewind_restarts_iteration
// origin: languages/php/tests/php/test_php_spl_min_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
$heap->insert(10);
$heap->insert(20);
$heap->rewind();
echo $heap->current() === 10 ? "REWIND_MIN_OK" : "FAIL";
