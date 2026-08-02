<?php
// vybe-test: php/spl/spl_priority_queue_extract_both
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_BOTH);
$pq->insert('low', 1);
$pq->insert('high', 9);
$pq->insert('mid', 5);
$item = $pq->extract();
echo $item['data'];
echo $item['priority'];
