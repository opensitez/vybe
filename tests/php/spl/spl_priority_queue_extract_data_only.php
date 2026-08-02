<?php
// vybe-test: php/spl/spl_priority_queue_extract_data_only
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_DATA);
$pq->insert('first', 1);
$pq->insert('second', 10);
echo $pq->current();
echo $pq->key();
