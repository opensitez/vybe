<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_bottom_peek_first_in
// origin: languages/php/tests/php/test_php_spl_queue_fifo_enqueue_dequeue.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->enqueue("job1");
$q->enqueue("job2");
echo $q->bottom() === "job1" ? "BOTTOM_JOB1_OK" : "FAIL";
