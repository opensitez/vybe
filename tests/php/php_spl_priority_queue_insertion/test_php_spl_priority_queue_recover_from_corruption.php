<?php
// vybe-test: php/php_spl_priority_queue_insertion/test_php_spl_priority_queue_recover_from_corruption
// origin: languages/php/tests/php/test_php_spl_priority_queue_insertion.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert("val", 1);
if (method_exists($pq, "recoverFromCorruption")) {
    $pq->recoverFromCorruption();
}
echo "RECOVER_OK";
