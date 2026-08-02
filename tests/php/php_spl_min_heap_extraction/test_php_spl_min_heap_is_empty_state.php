<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_is_empty_state
// origin: languages/php/tests/php/test_php_spl_min_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
echo $heap->isEmpty() ? "IS_EMPTY_TRUE" : "FAIL";
$heap->insert(1);
echo !$heap->isEmpty() ? " IS_EMPTY_FALSE" : " FAIL";
