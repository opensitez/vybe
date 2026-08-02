<?php
// vybe-test: php/php_spl_min_heap_extraction/test_php_spl_min_heap_extract_empty_throws_runtime_exception
// origin: languages/php/tests/php/test_php_spl_min_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMinHeap();
try {
    $heap->extract();
} catch (RuntimeException $e) {
    echo "MIN_HEAP_EMPTY_EX";
}
