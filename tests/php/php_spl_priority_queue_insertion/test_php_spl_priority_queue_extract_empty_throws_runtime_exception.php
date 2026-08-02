<?php
// vybe-test: php/php_spl_priority_queue_insertion/test_php_spl_priority_queue_extract_empty_throws_runtime_exception
// origin: languages/php/tests/php/test_php_spl_priority_queue_insertion.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
try {
    $pq->extract();
} catch (RuntimeException $e) {
    echo "PRIORITY_EMPTY_EX";
}
