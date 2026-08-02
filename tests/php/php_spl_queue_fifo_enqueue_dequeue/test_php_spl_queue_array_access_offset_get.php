<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_array_access_offset_get
// origin: languages/php/tests/php/test_php_spl_queue_fifo_enqueue_dequeue.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->enqueue("x");
$q->enqueue("y");
echo $q[0] === "x" && $q[1] === "y" ? "FIFO_INDEX_OK" : "FAIL";
