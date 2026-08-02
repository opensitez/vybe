<?php
// vybe-test: php/php_spl_priority_queue_insertion/test_php_spl_priority_queue_top_peeks_element
// origin: languages/php/tests/php/test_php_spl_priority_queue_insertion.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert("item", 5);
echo $pq->top() === "item" && count($pq) === 1 ? "PEEK_OK" : "FAIL";
