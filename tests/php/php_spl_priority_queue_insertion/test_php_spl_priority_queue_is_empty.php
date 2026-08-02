<?php
// vybe-test: php/php_spl_priority_queue_insertion/test_php_spl_priority_queue_is_empty
// origin: languages/php/tests/php/test_php_spl_priority_queue_insertion.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
echo $pq->isEmpty() ? "EMPTY" : "NOT_EMPTY";
