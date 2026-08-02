<?php
// vybe-test: php/php_spl_max_heap_extraction/test_php_spl_max_heap_is_empty
// origin: languages/php/tests/php/test_php_spl_max_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMaxHeap();
echo $heap->isEmpty() ? "EMPTY" : "NOT_EMPTY";
