<?php
// vybe-test: php/spl_extra/spl_priority_queue_priority_order
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert('low',    1);
$pq->insert('high',   10);
$pq->insert('medium', 5);
$out = [];
while (!$pq->isEmpty()) { $out[] = $pq->extract(); }
echo implode(',', $out); // high,medium,low
