<?php
// vybe-test: php/spl/spl_priority_queue_count
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert('a', 1);
$pq->insert('b', 2);
$pq->insert('c', 3);
echo $pq->count();
