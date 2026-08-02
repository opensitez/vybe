<?php
// vybe-test: php/spl/spl_priority_queue_basic
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert('low task',    1);
$pq->insert('high task',   10);
$pq->insert('medium task', 5);
$result = [];
while (!$pq->isEmpty()) { $result[] = $pq->extract(); }
echo implode(',', $result);
