<?php
// vybe-test: php/php_spl_priority_queue_insertion/test_php_spl_priority_queue_same_priority_fifo_order
// origin: languages/php/tests/php/test_php_spl_priority_queue_insertion.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert("first", 10);
$pq->insert("second", 10);
$e1 = $pq->extract();
echo ($e1 === "first" || $e1 === "second") ? "SAME_PRIORITY_EXTRACT_OK" : "FAIL";
