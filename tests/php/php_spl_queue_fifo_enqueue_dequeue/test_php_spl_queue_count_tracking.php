<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_count_tracking
// origin: languages/php/tests/php/test_php_spl_queue_fifo_enqueue_dequeue.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->enqueue(1);
$q->enqueue(2);
$q->dequeue();
echo count($q) === 1 ? "COUNT_1_OK" : "FAIL";
