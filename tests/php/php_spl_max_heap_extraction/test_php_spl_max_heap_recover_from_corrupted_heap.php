<?php
// vybe-test: php/php_spl_max_heap_extraction/test_php_spl_max_heap_recover_from_corrupted_heap
// origin: languages/php/tests/php/test_php_spl_max_heap_extraction.rs
// vybe-test-mode: compile

$heap = new SplMaxHeap();
$heap->insert(1);
if (method_exists($heap, "recoverFromCorruption")) {
    $heap->recoverFromCorruption();
}
echo "RECOVER_OK";
