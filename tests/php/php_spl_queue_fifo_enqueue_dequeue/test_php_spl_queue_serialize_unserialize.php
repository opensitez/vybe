<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_serialize_unserialize
// origin: languages/php/tests/php/test_php_spl_queue_fifo_enqueue_dequeue.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->enqueue("task_alpha");
$s = serialize($q);
$restored = unserialize($s);
echo $restored->dequeue() === "task_alpha" ? "SERIALIZE_QUEUE_OK" : "FAIL";
