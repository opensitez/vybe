<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_is_empty_check
// origin: languages/php/tests/php/test_php_spl_queue_fifo_enqueue_dequeue.rs
// vybe-test-mode: compile

$q = new SplQueue();
echo $q->isEmpty() ? "EMPTY_OK" : "FAIL";
