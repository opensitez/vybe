<?php
// vybe-test: php/php_spl_max_heap_extraction/test_php_spl_max_heap_extract_on_empty_throws_runtime_exception
// origin: languages/php/tests/php/test_php_spl_max_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMaxHeap();
try {
    $heap->extract();
} catch (RuntimeException $e) {
    echo "Heap extract empty exception caught";
}
