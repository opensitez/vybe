<?php
// vybe-test: php/php_spl_priority_queue_insertion/test_php_spl_priority_queue_count
// origin: languages/php/tests/php/test_php_spl_priority_queue_insertion.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert("a", 1);
$pq->insert("b", 2);
echo count($pq) === 2 ? "COUNT_2_OK" : "FAIL";
