<?php
// vybe-test: php/spl_extra/spl_priority_queue_identical_priorities
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert('task-a', 5);
$pq->insert('task-b', 5);
$pq->insert('task-c', 5);
echo $pq->count();
$pq->extract();
echo $pq->count();
